"use strict";

const express = require("express");
const path = require("path");
const ExcelJS = require("exceljs");
const PDFDocument = require("pdfkit");

const app = express();
const PORT = process.env.PORT || 3000;

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------
const JIRA_SITE = process.env.JIRA_SITE || "accertetecnologia.atlassian.net";
const JIRA_EMAIL = process.env.JIRA_EMAIL || "";
const JIRA_API_TOKEN = process.env.JIRA_API_TOKEN || "";
const PROJECT_KEY = process.env.JIRA_PROJECT || "SUPORTE";
// 395 dias (~13 meses) dá folga pra cobrir o filtro "Últimos 12 meses" com
// alguma margem — sem isso, o período mais longo do filtro ficaria truncado
// pelos dados que o servidor nem buscou no Jira.
const WINDOW_DAYS = parseInt(process.env.WINDOW_DAYS || "395", 10);

const APP_USER = process.env.APP_USER || "accerte";
const APP_PASSWORD = process.env.APP_PASSWORD || "";
const CRON_TOKEN = process.env.CRON_TOKEN || "";

// Auto-refresh: o próprio processo busca dados novos no Jira periodicamente,
// sem depender de clique manual nem de um Cron Job separado (e sem custo extra).
const AUTO_REFRESH_MINUTES = parseInt(process.env.AUTO_REFRESH_MINUTES || "5", 10);

const ANALYSTS = [
  { slot: 1, name: "Lucas Viana Hahn", short: "Lucas V. Hahn", accountId: "712020:081c9569-f2f3-4c30-914c-507a0def029b" },
  { slot: 2, name: "Vinícius Felipe de Souza Soares", short: "Vinícius Soares", accountId: "712020:8148d69f-4c1a-41d2-8d15-e1780d575fff" },
  { slot: 3, name: "Alexsander", short: "Alexsander", accountId: "712020:18affbde-27f1-436c-95d0-64902fa1d97e" },
  { slot: 4, name: "Lucas Alexandre Brandão Lopes", short: "Lucas A. Brandão", accountId: "712020:16eecd49-9135-471c-bc9d-1008d6c66909" },
];
const ACCOUNT_TO_SLOT = {};
ANALYSTS.forEach(function (a) { ACCOUNT_TO_SLOT[a.accountId] = a.slot; });

// ---------------------------------------------------------------------------
// In-memory state
// ---------------------------------------------------------------------------
let cache = { tickets: [], meta: null };
let refreshState = { running: false, startedAt: null, finishedAt: null, error: null, progress: 0 };

// ---------------------------------------------------------------------------
// Jira REST API v3 (email + API token, Basic Auth)
// ---------------------------------------------------------------------------
function jiraAuthHeader() {
  const raw = JIRA_EMAIL + ":" + JIRA_API_TOKEN;
  return "Basic " + Buffer.from(raw, "utf8").toString("base64");
}

async function jiraSearchPage(jql, fields, nextPageToken) {
  const body = { jql: jql, fields: fields, maxResults: 100 };
  if (nextPageToken) body.nextPageToken = nextPageToken;

  const resp = await fetch("https://" + JIRA_SITE + "/rest/api/3/search/jql", {
    method: "POST",
    headers: {
      Authorization: jiraAuthHeader(),
      "Content-Type": "application/json",
      Accept: "application/json",
    },
    body: JSON.stringify(body),
  });

  if (!resp.ok) {
    let text = "";
    try { text = await resp.text(); } catch (e) { /* ignore */ }
    throw new Error("Jira respondeu HTTP " + resp.status + (text ? (": " + text.slice(0, 300)) : ""));
  }
  return resp.json();
}

function issueToTuple(issue, slot) {
  const key = issue.key;
  const summary = String((issue.fields && issue.fields.summary) || "").slice(0, 220);
  const createdMs = new Date(issue.fields.created).getTime();
  const priority = (issue.fields && issue.fields.priority && issue.fields.priority.name) || null;

  let elapsedMin = null;
  let breached = null;
  let goalMin = null;
  let ongoingBreachMs = null;

  const sla = issue.fields && issue.fields.customfield_10121;
  if (sla && Array.isArray(sla.completedCycles) && sla.completedCycles.length > 0) {
    const cycle = sla.completedCycles[0];
    if (cycle.elapsedTime && typeof cycle.elapsedTime.millis === "number") {
      elapsedMin = cycle.elapsedTime.millis / 60000;
    }
    if (cycle.breached === true) breached = 1;
    else if (cycle.breached === false) breached = 0;
    if (cycle.goalDuration && typeof cycle.goalDuration.millis === "number") {
      goalMin = cycle.goalDuration.millis / 60000;
    }
  } else if (sla && sla.ongoingCycle) {
    // No first response yet: the SLA clock is still running. Capture the goal
    // and the exact wall-clock breach moment (Jira already accounts for
    // calendar/business-hours pauses in breachTime) so the SLA watch tab can
    // show tickets counting down to a 30-minute breach in real time.
    const oc = sla.ongoingCycle;
    if (oc.goalDuration && typeof oc.goalDuration.millis === "number") {
      goalMin = oc.goalDuration.millis / 60000;
    }
    if (oc.breachTime && typeof oc.breachTime.epochMillis === "number") {
      ongoingBreachMs = oc.breachTime.epochMillis;
    }
  }

  return [slot, createdMs, elapsedMin, breached, key, summary, goalMin, ongoingBreachMs, priority];
}

async function pullAllTickets(onProgress) {
  if (!JIRA_EMAIL || !JIRA_API_TOKEN) {
    throw new Error("JIRA_EMAIL / JIRA_API_TOKEN não configurados no servidor.");
  }

  const accountIds = ANALYSTS.map(function (a) { return '"' + a.accountId + '"'; }).join(",");
  const jql =
    "project = " + PROJECT_KEY +
    " AND assignee in (" + accountIds + ")" +
    " AND created >= -" + WINDOW_DAYS + "d" +
    " ORDER BY created DESC";
  const fields = ["summary", "created", "customfield_10121", "assignee", "priority"];

  const tickets = [];
  let nextPageToken;
  const MAX_PAGES = 400;

  for (let page = 0; page < MAX_PAGES; page++) {
    const data = await jiraSearchPage(jql, fields, nextPageToken);
    const issues = data.issues || [];

    for (const issue of issues) {
      const accountId = issue.fields && issue.fields.assignee && issue.fields.assignee.accountId;
      const slot = ACCOUNT_TO_SLOT[accountId];
      if (!slot) continue;
      tickets.push(issueToTuple(issue, slot));
    }

    if (onProgress) onProgress(tickets.length);

    const isLast = data.isLast === true || issues.length === 0 || !data.nextPageToken;
    if (isLast) break;
    nextPageToken = data.nextPageToken;
  }

  return tickets;
}

async function doRefresh() {
  if (refreshState.running) return;
  refreshState = { running: true, startedAt: Date.now(), finishedAt: null, error: null, progress: 0 };
  console.log("[refresh] iniciado");
  try {
    const tickets = await pullAllTickets(function (count) {
      refreshState.progress = count;
    });
    cache = {
      tickets: tickets,
      meta: {
        generated_at: new Date().toISOString(),
        source_project: PROJECT_KEY,
        site_host: JIRA_SITE,
        window_days: WINDOW_DAYS,
      },
    };
    refreshState.finishedAt = Date.now();
    console.log("[refresh] concluído: " + tickets.length + " chamados");
  } catch (err) {
    refreshState.error = (err && err.message) || String(err);
    refreshState.finishedAt = Date.now();
    console.error("[refresh] falhou:", refreshState.error);
  } finally {
    refreshState.running = false;
  }
}

// ---------------------------------------------------------------------------
// Auth
// ---------------------------------------------------------------------------
function checkBasicAuth(req) {
  if (!APP_PASSWORD) return false;
  const header = req.headers.authorization || "";
  const parts = header.split(" ");
  if (parts.length !== 2 || parts[0] !== "Basic") return false;
  let decoded;
  try {
    decoded = Buffer.from(parts[1], "base64").toString("utf8");
  } catch (e) {
    return false;
  }
  const idx = decoded.indexOf(":");
  if (idx === -1) return false;
  const user = decoded.slice(0, idx);
  const pass = decoded.slice(idx + 1);
  return user === APP_USER && pass === APP_PASSWORD;
}

function requireAuth(req, res, next) {
  if (checkBasicAuth(req)) return next();
  res.set("WWW-Authenticate", 'Basic realm="TPA Dashboard"');
  res.status(401).send("Autenticação necessária.");
}

// ---------------------------------------------------------------------------
// Routes
// ---------------------------------------------------------------------------
app.use(express.json());

app.get("/healthz", function (req, res) {
  res.send("ok");
});

app.get("/favicon.ico", function (req, res) {
  res.status(204).end();
});

// Refresh can be triggered either by an authenticated browser session, or by
// the scheduled cron job presenting the shared CRON_TOKEN header.
app.post("/api/refresh", function (req, res) {
  const cronHeader = req.headers["x-cron-token"];
  const authorized = (CRON_TOKEN && cronHeader === CRON_TOKEN) || checkBasicAuth(req);
  if (!authorized) {
    res.set("WWW-Authenticate", 'Basic realm="TPA Dashboard"');
    return res.status(401).json({ error: "não autorizado" });
  }
  if (refreshState.running) {
    return res.json({ started: false, alreadyRunning: true });
  }
  doRefresh();
  res.json({ started: true });
});

app.use(requireAuth);

app.get("/api/tickets", function (req, res) {
  res.json({
    meta: cache.meta,
    analysts: ANALYSTS.map(function (a) { return { slot: a.slot, name: a.name, short: a.short }; }),
    tickets: cache.tickets,
  });
});

app.get("/api/status", function (req, res) {
  res.json({
    running: refreshState.running,
    startedAt: refreshState.startedAt,
    finishedAt: refreshState.finishedAt,
    error: refreshState.error,
    progress: refreshState.progress,
    hasData: !!cache.meta,
    generated_at: cache.meta ? cache.meta.generated_at : null,
  });
});

// Gera uma planilha .xlsx formatada a partir de linhas já filtradas/ordenadas
// no navegador (drill-downs de pendentes/estouro de SLA), para manter a
// exportação sempre igual ao que está sendo exibido na tela.
app.post("/api/export/xlsx", async function (req, res) {
  try {
    var body = req.body || {};
    var title = String(body.title || "Exportação").slice(0, 90);
    var columns = Array.isArray(body.columns) ? body.columns : [];
    var rows = Array.isArray(body.rows) ? body.rows : [];
    var filename = String(body.filename || "export.xlsx").replace(/[^a-zA-Z0-9_\-.]/g, "_");
    if (!columns.length) {
      return res.status(400).json({ error: "nenhuma coluna informada" });
    }
    if (rows.length > 20000) {
      return res.status(400).json({ error: "muitas linhas para exportar" });
    }

    var workbook = new ExcelJS.Workbook();
    workbook.creator = "TPA Dashboard - Central de Suporte Accerte";
    workbook.created = new Date();

    var sheetName = title.replace(/[\\/*?:[\]]/g, " ").slice(0, 31) || "Dados";
    var sheet = workbook.addWorksheet(sheetName, {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    sheet.columns = columns.map(function (c) {
      return { header: String(c.header || ""), key: String(c.key || c.header || ""), width: c.width || 20 };
    });

    rows.forEach(function (r) {
      sheet.addRow(Array.isArray(r) ? r : columns.map(function (c) { return r[c.key]; }));
    });

    var headerRow = sheet.getRow(1);
    headerRow.eachCell(function (cell) {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F2937" } };
      cell.alignment = { vertical: "middle", horizontal: "left" };
      cell.border = { bottom: { style: "thin", color: { argb: "FF000000" } } };
    });
    headerRow.height = 20;

    for (var i = 2; i <= sheet.rowCount; i++) {
      var row = sheet.getRow(i);
      row.eachCell(function (cell) {
        cell.border = { bottom: { style: "thin", color: { argb: "FFE5E7EB" } } };
        cell.alignment = { vertical: "middle" };
      });
      if (i % 2 === 0) {
        row.eachCell(function (cell) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF9FAFB" } };
        });
      }
    }

    if (columns.length) {
      sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: columns.length } };
    }

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="' + filename + '"');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error("[export] falhou:", (err && err.message) || err);
    if (!res.headersSent) {
      res.status(500).json({ error: (err && err.message) || String(err) });
    } else {
      res.end();
    }
  }
});

// ---------------------------------------------------------------------------
// Relatório em PDF — layout desenhado com pdfkit a partir de um payload já
// calculado no navegador (mesmos números exibidos na tela, pra não haver
// divergência entre o que o usuário vê e o que sai no relatório).
// ---------------------------------------------------------------------------
const PDF_COLORS = {
  header: "#1F2937",
  headerText: "#FFFFFF",
  text: "#0b0b0b",
  muted: "#6b7280",
  border: "#e5e7eb",
  rowAlt: "#f9fafb",
  accent: "#2a78d6",
  good: "#0ca30c",
  warning: "#fab219",
  serious: "#ec835a",
  critical: "#d03b3b",
};

function pdfEnsureSpace(doc, y, needed, marginTop, marginBottom) {
  var pageBottom = doc.page.height - marginBottom;
  if (y + needed > pageBottom) {
    doc.addPage();
    return marginTop;
  }
  return y;
}

function pdfSectionTitle(doc, text, x, y, width) {
  doc.font("Helvetica-Bold").fontSize(12.5).fillColor(PDF_COLORS.text);
  doc.text(String(text || ""), x, y, { width: width });
  return y + 20;
}

function pdfKpiRow(doc, kpis, x, y, width) {
  var gap = 10;
  var n = kpis.length || 1;
  var boxW = (width - gap * (n - 1)) / n;
  var boxH = 58;
  kpis.forEach(function (k, i) {
    var bx = x + i * (boxW + gap);
    doc.roundedRect(bx, y, boxW, boxH, 4).fillAndStroke("#ffffff", PDF_COLORS.border);
    doc.font("Helvetica").fontSize(8.5).fillColor(PDF_COLORS.muted);
    doc.text(String(k.label || "").toUpperCase(), bx + 10, y + 9, { width: boxW - 20 });
    doc.font("Helvetica-Bold").fontSize(16).fillColor(PDF_COLORS.text);
    doc.text(String(k.value || ""), bx + 10, y + 22, { width: boxW - 20 });
    if (k.sub) {
      doc.font("Helvetica").fontSize(7.5).fillColor(PDF_COLORS.muted);
      doc.text(String(k.sub), bx + 10, y + 42, { width: boxW - 20, ellipsis: true });
    }
  });
  return y + boxH + 18;
}

function pdfTable(doc, columns, rows, x, y, width, marginTop, marginBottom) {
  var rowH = 20;
  var headH = 22;
  var widths = columns.map(function (c) { return c.width || width / columns.length; });
  var totalW = widths.reduce(function (s, w) { return s + w; }, 0);
  if (totalW !== width) {
    var scale = width / totalW;
    widths = widths.map(function (w) { return w * scale; });
  }
  y = pdfEnsureSpace(doc, y, headH + rowH, marginTop, marginBottom);
  doc.rect(x, y, width, headH).fill(PDF_COLORS.header);
  var cx = x;
  columns.forEach(function (c, i) {
    doc.font("Helvetica-Bold").fontSize(9).fillColor(PDF_COLORS.headerText);
    doc.text(String(c.header || ""), cx + 8, y + 6.5, { width: widths[i] - 12, align: c.align || "left" });
    cx += widths[i];
  });
  y += headH;
  rows.forEach(function (r, ri) {
    y = pdfEnsureSpace(doc, y, rowH, marginTop, marginBottom);
    if (ri % 2 === 1) doc.rect(x, y, width, rowH).fill(PDF_COLORS.rowAlt);
    var cx2 = x;
    r.forEach(function (val, ci) {
      doc.font("Helvetica").fontSize(9.5).fillColor(PDF_COLORS.text);
      doc.text(String(val == null ? "—" : val), cx2 + 8, y + 5.5, {
        width: widths[ci] - 12,
        align: (columns[ci] && columns[ci].align) || "left",
      });
      cx2 += widths[ci];
    });
    doc.moveTo(x, y + rowH).lineTo(x + width, y + rowH).strokeColor(PDF_COLORS.border).lineWidth(0.5).stroke();
    y += rowH;
  });
  return y + 16;
}

function pdfBarList(doc, points, x, y, width, opts) {
  opts = opts || {};
  var rowH = 17;
  var labelW = 46;
  var valueW = 56;
  var trackX = x + labelW + 6;
  var trackW = width - labelW - valueW - 12;
  var maxVal = Math.max.apply(
    null,
    points.map(function (p) { return p.value == null ? 0 : p.value; }).concat([opts.maxHint || 1])
  );
  points.forEach(function (p) {
    doc.font("Helvetica").fontSize(8.5).fillColor(PDF_COLORS.muted);
    doc.text(String(p.label || ""), x, y + 4, { width: labelW, align: "left" });
    doc.rect(trackX, y + 2, trackW, 11).fill(PDF_COLORS.border);
    if (p.value != null && maxVal > 0) {
      var w = Math.max(3, (p.value / maxVal) * trackW);
      doc.rect(trackX, y + 2, w, 11).fill(opts.barColor || PDF_COLORS.accent);
    }
    doc.font("Helvetica-Bold").fontSize(8.5).fillColor(PDF_COLORS.text);
    var valTxt = p.value == null ? "—" : (Math.round(p.value * 10) / 10).toLocaleString("pt-BR") + (opts.unit || "");
    doc.text(valTxt, trackX + trackW + 8, y + 4, { width: valueW - 8, align: "left" });
    y += rowH;
  });
  return y + 6;
}

function pdfSummaryBox(doc, text, tone, x, y, width) {
  if (!text) return y;
  var bg = tone === "good" ? "#e7f7e7" : tone === "bad" ? "#fbeaea" : "#f3f4f6";
  var fg = tone === "good" ? PDF_COLORS.good : tone === "bad" ? PDF_COLORS.critical : PDF_COLORS.muted;
  var h = doc.heightOfString(text, { width: width - 20, fontSize: 9.5 }) + 16;
  doc.roundedRect(x, y, width, h, 4).fill(bg);
  doc.font("Helvetica").fontSize(9.5).fillColor(fg);
  doc.text(text, x + 10, y + 8, { width: width - 20 });
  return y + h + 18;
}

function buildPdfReport(doc, payload) {
  var margin = 40;
  var pageW = doc.page.width;
  var contentW = pageW - margin * 2;

  doc.rect(0, 0, pageW, 78).fill(PDF_COLORS.header);
  doc.font("Helvetica-Bold").fontSize(17).fillColor(PDF_COLORS.headerText);
  doc.text(String(payload.title || "Relatório de Desempenho — TPA"), margin, 20, { width: contentW - 160 });
  doc.font("Helvetica").fontSize(10).fillColor("#c9d3de");
  doc.text(String(payload.scopeLabel || ""), margin, 44, { width: contentW - 160 });
  doc.font("Helvetica").fontSize(8).fillColor("#c9d3de");
  doc.text("Gerado em " + String(payload.generatedAt || ""), margin, 60, { width: contentW - 160 });
  doc.font("Helvetica-Bold").fontSize(9).fillColor(PDF_COLORS.headerText);
  doc.text("Central de Suporte Accerte", margin, 20, { width: contentW, align: "right" });

  var y = 96;

  if (Array.isArray(payload.kpis) && payload.kpis.length) {
    y = pdfEnsureSpace(doc, y, 76, 40, 40);
    y = pdfKpiRow(doc, payload.kpis, margin, y, contentW);
  }

  if (payload.ranking && Array.isArray(payload.ranking.rows) && payload.ranking.rows.length) {
    y = pdfEnsureSpace(doc, y, 40, 40, 40);
    y = pdfSectionTitle(doc, "Ranking por analista", margin, y, contentW);
    y = pdfTable(doc, payload.ranking.columns, payload.ranking.rows, margin, y, contentW, 40, 40);
  }

  if (payload.evolution && Array.isArray(payload.evolution.points) && payload.evolution.points.length) {
    y = pdfEnsureSpace(doc, y, 40, 40, 40);
    y = pdfSectionTitle(doc, payload.evolution.title || "Evolução do TPA (últimos 6 meses)", margin, y, contentW);
    y = pdfBarList(doc, payload.evolution.points, margin, y, contentW, {
      unit: " min",
      barColor: PDF_COLORS.accent,
    });
    y = pdfSummaryBox(doc, payload.evolution.summary, payload.evolution.trend, margin, y, contentW);
  }

  if (payload.compliance && Array.isArray(payload.compliance.points) && payload.compliance.points.length) {
    y = pdfEnsureSpace(doc, y, 40, 40, 40);
    y = pdfSectionTitle(doc, payload.compliance.title || "Cumprimento da Meta (últimos 6 meses)", margin, y, contentW);
    y = pdfBarList(doc, payload.compliance.points, margin, y, contentW, {
      unit: "%",
      barColor: PDF_COLORS.good,
      maxHint: 100,
    });
    y = pdfSummaryBox(doc, payload.compliance.summary, payload.compliance.trend, margin, y, contentW);
  }

  var footerY = doc.page.height - 34;
  doc.font("Helvetica").fontSize(7.5).fillColor(PDF_COLORS.muted);
  doc.text(
    "TPA = tempo até a primeira resposta (SLA do Jira Service Management, projeto " +
      String(payload.projectKey || "SUPORTE") +
      "). Chamados pendentes de resposta não entram na média nem na taxa de cumprimento.",
    margin,
    footerY,
    { width: contentW }
  );
}

app.post("/api/export/pdf", function (req, res) {
  try {
    var payload = req.body || {};
    var filename = String(payload.filename || "relatorio_tpa.pdf").replace(/[^a-zA-Z0-9_\-.]/g, "_");
    var doc = new PDFDocument({
      size: "A4",
      margins: { top: 40, bottom: 15, left: 40, right: 40 },
      bufferPages: true,
    });
    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", 'attachment; filename="' + filename + '"');
    doc.pipe(res);
    buildPdfReport(doc, payload);
    doc.end();
  } catch (err) {
    console.error("[export-pdf] falhou:", (err && err.message) || err);
    if (!res.headersSent) {
      res.status(500).json({ error: (err && err.message) || String(err) });
    } else {
      res.end();
    }
  }
});

app.get(["/", "/index.html"], function (req, res) {
  res.sendFile(path.join(__dirname, "index.html"));
});

app.listen(PORT, function () {
  console.log("TPA dashboard ouvindo na porta " + PORT);
  // Warm the cache on boot so the first visitor doesn't hit an empty dataset.
  doRefresh();

  // Mantém os dados atualizados automaticamente enquanto o processo estiver
  // no ar. No plano free do Render o serviço "dorme" após ficar sem receber
  // requisições por um tempo; quando ele acorda, o doRefresh() do boot acima
  // já cuida de trazer dados novos, e este intervalo retoma normalmente.
  if (AUTO_REFRESH_MINUTES > 0) {
    console.log("[refresh] auto-refresh ativado a cada " + AUTO_REFRESH_MINUTES + " min");
    setInterval(function () {
      console.log("[refresh] disparo automático (" + AUTO_REFRESH_MINUTES + " min)");
      doRefresh();
    }, AUTO_REFRESH_MINUTES * 60 * 1000);
  }
});
