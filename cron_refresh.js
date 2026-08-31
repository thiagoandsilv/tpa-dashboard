"use strict";

// Runs on Render's scheduled Cron Job. Hits the web service's /api/refresh
// endpoint with the shared secret header, so the dashboard's data is
// refreshed automatically every day without anyone needing to click the
// "Atualizar" button.

const SERVICE_URL = process.env.SERVICE_URL || "";
const CRON_TOKEN = process.env.CRON_TOKEN || "";

async function main() {
  if (!SERVICE_URL) throw new Error("SERVICE_URL não configurado");
  if (!CRON_TOKEN) throw new Error("CRON_TOKEN não configurado");

  const url = SERVICE_URL.replace(/\/+$/, "") + "/api/refresh";
  console.log("[cron] disparando atualização em " + url);

  const resp = await fetch(url, {
    method: "POST",
    headers: { "x-cron-token": CRON_TOKEN },
  });

  const text = await resp.text().catch(function () { return ""; });
  console.log("[cron] HTTP " + resp.status + " " + text.slice(0, 300));

  if (!resp.ok) {
    process.exitCode = 1;
  }
}

main().catch(function (err) {
  console.error("[cron] falhou:", err && err.message ? err.message : err);
  process.exitCode = 1;
});
