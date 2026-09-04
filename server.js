"use strict";

const express = require("express");
const path = require("path");
const ExcelJS = require("exceljs");

const app = express();
const PORT = process.env.PORT || 3000;

// ---------------------------------------------------------------------------
// Config
// ---------------------------------------------------------------------------
const JIRA_SITE = process.env.JIRA_SITE || "accertetecnologia.atlassian.net";
const JIRA_EMAIL = process.env.JIRA_EMAIL || "";
const JIRA_API_TOKEN = process.env.JIRA_API_TOKEN || "";
const PROJECT_KEY = process.env.JIRA_PROJECT || "SUPORTE";
const WINDOW_DAYS = parseInt(process.env.WINDOW_DAYS || "185", 10);

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

  return [slot, createdMs, elapsedMin, breached, key, summary, goalMin, ongoingBreachMs];
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
  const fields = ["summary", "created", "customfield_10121", "assignee"];

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
