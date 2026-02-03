import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

/**
 * Deterministic /api/chat
 * - NO OpenAI calls.
 * - Answers ONLY using the state.json data in Azure Blob Storage.
 * - Supports "TEMPLATE:<id>\nPARAMS:<json>" prompts emitted by ChatOverlay quick questions.
 *
 * Expected storage layout (defaults can be overridden with env vars):
 *   container: workback
 *   state blob: state.json
 *   backups weekly: backups/weekly/<anything with YYYY-MM-DD>.json
 *
 * Env vars:
 *   BLOB_CONNECTION_STRING (or AzureWebJobsStorage)
 *   STATE_CONTAINER (default "workback")
 *   STATE_BLOB (default "state.json")
 *   BACKUP_PREFIX (default "backups/weekly/")
 *   BACKUP_MAX_BYTES (default 8_000_000)
 */

function safeStr(v) {
  return (v ?? "").toString().trim();
}

function isoToUtcDate(s) {
  if (!s) return null;
  const m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  const y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  const dt = new Date(Date.UTC(y, mo - 1, d));
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function addDaysUTC(iso, days) {
  const dt = isoToUtcDate(iso);
  if (!dt || typeof days !== "number") return null;
  const out = new Date(dt.getTime() + days * 86400000);
  return out.toISOString().slice(0, 10);
}

function diffDaysUTC(aISO, bISO) {
  const a = isoToUtcDate(aISO);
  const b = isoToUtcDate(bISO);
  if (!a || !b) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

function YESNO(b) {
  return b ? "YES" : "NO";
}

function getBlobClients() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  if (!conn) throw new Error("Storage connection string is not set (BLOB_CONNECTION_STRING or AzureWebJobsStorage)");
  const service = BlobServiceClient.fromConnectionString(conn);
  return { container: service.getContainerClient(containerName) };
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

async function readJsonBlob(container, blobName, maxBytes = 8_000_000) {
  const blob = container.getBlobClient(blobName);
  const props = await blob.getProperties().catch(() => null);
  if (props && props.contentLength && props.contentLength > maxBytes) {
    throw new Error(`Blob too large (${props.contentLength} bytes) for deterministic function limit`);
  }
  const dl = await blob.download();
  const text = await streamToString(dl.readableStreamBody);
  return text ? JSON.parse(text) : null;
}

/* ---------------- index build ---------------- */

function buildIndexFromState(state) {
  const projects = Array.isArray(state?.projects) ? state.projects : [];
  const globalA = Number(state?.globalDaysReqToStatusA ?? 28);
  const globalF = Number(state?.globalDaysStatusAToFirstIssue ?? 28);

  // programme levels
  const programme = projects.flatMap((p) =>
    (p.master || []).flatMap((m) =>
      (m.levels || []).map((lv) => ({
        project: safeStr(p.name),
        blockZone: safeStr(m.blockZone),
        level: safeStr(lv.name),
        startDate: safeStr(lv.startDate),
        finishDate: safeStr(lv.finishDate),
      }))
    )
  );

  // tracker rows
  const rows = [];
  projects.forEach((p) => {
    const responsibilities = Array.isArray(p.responsibilities) ? p.responsibilities : [];
    const pages = Array.isArray(p.pages) ? p.pages : [];

    const supplierByRespId = new Map(responsibilities.map((r) => [r.id, safeStr(r.supplier)]));

    pages
      .filter((pg) => !pg?.meta?.isMaster)
      .forEach((pg) => {
        const supplier = supplierByRespId.get(pg?.meta?.responsibilityId) || "—";
        (pg.rows || [])
          .filter((r) => r && r.kind !== "header")
          .filter((r) => !r.notRequired)
          .forEach((r) => {
            const anchorKey = safeStr(r.anchorKey) || "requiredOnSite";
            const anchor = safeStr(r.anchorDateISO);

            const daysA = Number(r.overrideDaysReqToStatusA ?? globalA);
            const daysF = Number(r.overrideDaysStatusAToFirstIssue ?? globalF);

            // Derive dates from anchor
            let requiredOnSite = null;
            let statusA = null;
            let firstIssue = null;

            if (anchorKey === "requiredOnSite") {
              requiredOnSite = anchor || null;
              statusA = requiredOnSite ? addDaysUTC(requiredOnSite, -daysA) : null;
              firstIssue = statusA ? addDaysUTC(statusA, daysF) : null;
            } else if (anchorKey === "statusA") {
              statusA = anchor || null;
              requiredOnSite = statusA ? addDaysUTC(statusA, daysA) : null;
              firstIssue = statusA ? addDaysUTC(statusA, daysF) : null;
            } else if (anchorKey === "firstIssue") {
              firstIssue = anchor || null;
              statusA = firstIssue ? addDaysUTC(firstIssue, -daysF) : null;
              requiredOnSite = statusA ? addDaysUTC(statusA, daysA) : null;
            } else {
              // fallback
              requiredOnSite = anchor || null;
              statusA = requiredOnSite ? addDaysUTC(requiredOnSite, -daysA) : null;
              firstIssue = statusA ? addDaysUTC(statusA, daysF) : null;
            }

            const comments = Array.isArray(r.comments) ? r.comments : [];
            const commentsText = comments
              .map((c) => [safeStr(c?.dateISO), safeStr(c?.name || c?.lockedBy), safeStr(c?.text)].filter(Boolean).join(" — "))
              .filter(Boolean)
              .join("\n");

            rows.push({
              key: [safeStr(p.name), safeStr(pg.name), safeStr(supplier), safeStr(r.item)].join(" | "),
              project: safeStr(p.name),
              responsibility: safeStr(pg.name),
              supplier,
              item: safeStr(r.item),
              requiredOnSite,
              statusA,
              firstIssue,
              ticks: {
                statusA: !!r.statusADone,
                firstIssue: !!r.firstIssueDone,
                completed: !!r.completed,
              },
              comments: commentsText,
              commentsCount: comments.length,
            });
          });
      });
  });

  return { programme, rows, projectNames: projects.map((p) => safeStr(p.name)).filter(Boolean) };
}

/* ---------------- filters + formatting ---------------- */

function withinRange(level, fromISO, toISO) {
  const s = isoToUtcDate(level?.startDate);
  const f = isoToUtcDate(level?.finishDate);
  const a = isoToUtcDate(fromISO);
  const b = isoToUtcDate(toISO);
  if (!s || !f || !a || !b) return false;
  // overlap inclusive
  return !(f.getTime() < a.getTime() || s.getTime() > b.getTime());
}

function formatRowsTable(rows, limit = 120) {
  const header = ["Project", "Responsibility", "Supplier", "Item", "Req On Site", "Status A", "First Issue", "Done?"].join(" | ");
  const sep = ["---","---","---","---","---","---","---","---"].join(" | ");
  const body = rows.slice(0, limit).map((r) => [
    r.project || "—",
    r.responsibility || "—",
    r.supplier || "—",
    r.item || "—",
    r.requiredOnSite || "—",
    r.statusA || "—",
    r.firstIssue || "—",
    YESNO(!!r.ticks?.completed),
  ].join(" | "));
  const more = rows.length > limit ? `\n\nShowing first ${limit} of ${rows.length}. Narrow by project to see more.` : "";
  return [header, sep, ...body].join("\n") + more;
}

function formatProgrammeTable(levels, limit = 200) {
  const header = ["Project", "Block/Zone", "Level", "Start", "Finish"].join(" | ");
  const sep = ["---","---","---","---","---"].join(" | ");
  const body = levels.slice(0, limit).map((p) => [
    p.project || "—",
    p.blockZone || "—",
    p.level || "—",
    p.startDate || "—",
    p.finishDate || "—",
  ].join(" | "));
  const more = levels.length > limit ? `\n\nShowing first ${limit} of ${levels.length}. Narrow by project to see more.` : "";
  return [header, sep, ...body].join("\n") + more;
}

function parseTemplateFromUserText(text) {
  const t = String(text || "");
  const m = t.match(/^\s*TEMPLATE:([a-z0-9_]+)\s*\n\s*PARAMS:(\{[\s\S]*\})\s*$/i);
  if (!m) return null;
  const id = m[1];
  let params = {};
  try {
    params = JSON.parse(m[2]);
  } catch {
    params = {};
  }
  return { id, params };
}

/* ---------------- template handlers ---------------- */

function handleTemplate({ id, params }, idx, todayISO) {
  const project = safeStr(params?.project);
  const dateFrom = safeStr(params?.dateFrom);
  const dateTo = safeStr(params?.dateTo);
  const asOfDate = safeStr(params?.asOfDate);

  const today = todayISO;

  const rows = idx.rows;
  const programme = idx.programme;

  const byProject = (r) => !project || safeStr(r.project).toLowerCase() === project.toLowerCase();

  const overdue = (iso) => {
    const d = isoToUtcDate(iso);
    const t = isoToUtcDate(today);
    if (!d || !t) return false;
    return d.getTime() < t.getTime();
  };

  if (id === "overdue_first_issue_all" || id === "overdue_first_issue_project") {
    const out = rows.filter(byProject).filter((r) => !r.ticks?.firstIssue).filter((r) => overdue(r.firstIssue));
    out.sort((a, b) => (safeStr(a.firstIssue) < safeStr(b.firstIssue) ? -1 : 1));
    return {
      title: `Overdue First Issue${project ? ` — ${project}` : ""} (today ${today})`,
      body: formatRowsTable(out),
    };
  }

  if (id === "overdue_statusA_all" || id === "overdue_statusA_project") {
    const out = rows.filter(byProject).filter((r) => !r.ticks?.statusA).filter((r) => overdue(r.statusA));
    out.sort((a, b) => (safeStr(a.statusA) < safeStr(b.statusA) ? -1 : 1));
    return {
      title: `Overdue Status A${project ? ` — ${project}` : ""} (today ${today})`,
      body: formatRowsTable(out),
    };
  }

  if (id === "statusA_approved_all" || id === "statusA_approved_project") {
    const out = rows.filter(byProject).filter((r) => !!r.ticks?.statusA);
    out.sort((a, b) => (safeStr(a.statusA) < safeStr(b.statusA) ? -1 : 1));
    return {
      title: `Status A Approved${project ? ` — ${project}` : ""}`,
      body: formatRowsTable(out),
    };
  }

  if (id === "done_all" || id === "done_project") {
    const out = rows.filter(byProject).filter((r) => !!r.ticks?.completed);
    out.sort((a, b) => (safeStr(a.requiredOnSite) < safeStr(b.requiredOnSite) ? -1 : 1));
    return {
      title: `Done items${project ? ` — ${project}` : ""}`,
      body: formatRowsTable(out),
    };
  }

  if (id === "programme_range_all" || id === "programme_range_project") {
    const out = programme
      .filter((p) => !project || safeStr(p.project).toLowerCase() === project.toLowerCase())
      .filter((p) => withinRange(p, dateFrom, dateTo));
    out.sort((a, b) => (safeStr(a.startDate) < safeStr(b.startDate) ? -1 : 1));
    return {
      title: `Programme on site ${dateFrom} → ${dateTo}${project ? ` — ${project}` : ""}`,
      body: formatProgrammeTable(out),
    };
  }

  if (id === "comments_all" || id === "comments_project") {
    const out = rows
      .filter(byProject)
      .filter((r) => (r.commentsCount || 0) > 0)
      .map((r) => ({
        project: r.project,
        responsibility: r.responsibility,
        supplier: r.supplier,
        item: r.item,
        comments: r.comments,
        commentsCount: r.commentsCount,
      }));

    // keep response readable
    const maxItems = 120;
    const lines = out.slice(0, maxItems).map((r) => {
      return `• ${r.project} | ${r.responsibility} | ${r.supplier} | ${r.item}\n${r.comments}`;
    });

    const more = out.length > maxItems ? `\n\nShowing first ${maxItems} of ${out.length}. Narrow by project to see more.` : "";
    return {
      title: `Comments${project ? ` — ${project}` : ""}`,
      body: lines.join("\n\n") + more,
    };
  }

  if (id === "compare_programme_all" || id === "compare_programme_project") {
    // This one is handled outside (needs a backup snapshot). Return a signal.
    return { needsBackupCompare: true, project, asOfDate };
  }

  return { title: "Unknown quick question", body: "That template id is not supported by the deterministic backend." };
}

/* ---------------- backup compare helpers ---------------- */

function extractAnyDateFromName(name) {
  const m = String(name || "").match(/(\d{4}-\d{2}-\d{2})/);
  return m ? m[1] : null;
}

function indexProgramme(programme) {
  const map = new Map();
  programme.forEach((p) => {
    const k = [safeStr(p.project), safeStr(p.blockZone), safeStr(p.level)].join(" | ");
    map.set(k, p);
  });
  return map;
}

function diffProgramme(cur, prev, projectFilter) {
  const curMap = indexProgramme(cur);
  const prevMap = indexProgramme(prev);
  const changes = [];

  for (const [k, c] of curMap.entries()) {
    if (projectFilter && safeStr(c.project).toLowerCase() !== projectFilter.toLowerCase()) continue;
    const p = prevMap.get(k);
    if (!p) continue;
    if (safeStr(c.startDate) !== safeStr(p.startDate) || safeStr(c.finishDate) !== safeStr(p.finishDate)) {
      changes.push({
        project: c.project,
        blockZone: c.blockZone,
        level: c.level,
        start: `${safeStr(p.startDate) || "—"} → ${safeStr(c.startDate) || "—"}`,
        finish: `${safeStr(p.finishDate) || "—"} → ${safeStr(c.finishDate) || "—"}`,
      });
    }
  }

  changes.sort((a, b) => (safeStr(a.project) + safeStr(a.blockZone) + safeStr(a.level)).localeCompare(safeStr(b.project) + safeStr(b.blockZone) + safeStr(b.level)));
  return changes;
}

function formatProgrammeChanges(changes, limit = 200) {
  const header = ["Project", "Block/Zone", "Level", "Start (before → after)", "Finish (before → after)"].join(" | ");
  const sep = ["---","---","---","---","---"].join(" | ");
  const body = changes.slice(0, limit).map((c) => [c.project, c.blockZone, c.level, c.start, c.finish].join(" | "));
  const more = changes.length > limit ? `\n\nShowing first ${limit} of ${changes.length}. Narrow by project to see more.` : "";
  return [header, sep, ...body].join("\n") + more;
}

/* ---------------- main handler ---------------- */

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const body = await req.json().catch(() => ({}));
      const messages = Array.isArray(body?.messages) ? body.messages : [];
      const appContext = body?.context || {};
      const userText = safeStr(messages[messages.length - 1]?.content);

      const todayISO = safeStr(appContext?.today) || new Date().toISOString().slice(0, 10);

      const { container } = getBlobClients();
      const stateBlob = process.env.STATE_BLOB || "state.json";
      const state = await readJsonBlob(container, stateBlob);
      const idx = buildIndexFromState(state || {});

      // Template-based deterministic quick questions
      const tpl = parseTemplateFromUserText(userText);
      if (tpl) {
        const result = handleTemplate(tpl, idx, todayISO);

        if (result?.needsBackupCompare) {
          const prefix = process.env.BACKUP_PREFIX || "backups/weekly/";
          const asOfDate = safeStr(result.asOfDate);
          if (!asOfDate) {
            return { status: 200, jsonBody: { answer: "Please provide an as-of date (YYYY-MM-DD) to compare against current." } };
          }

          // Find the newest backup <= asOfDate
          const candidates = [];
          for await (const blob of container.listBlobsFlat({ prefix })) {
            const d = extractAnyDateFromName(blob.name);
            if (!d) continue;
            if (d <= asOfDate) candidates.push({ name: blob.name, date: d });
          }
          candidates.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));

          if (!candidates.length) {
            return {
              status: 200,
              jsonBody: { answer: `No weekly backups found on or before ${asOfDate} under prefix "${prefix}".` },
            };
          }

          const chosen = candidates[0];
          const maxBytes = Number(process.env.BACKUP_MAX_BYTES || 8_000_000);
          const backupState = await readJsonBlob(container, chosen.name, maxBytes);
          const backupIdx = buildIndexFromState(backupState || {});

          const changes = diffProgramme(idx.programme, backupIdx.programme, result.project);
          const title = `Programme changes (${chosen.date} → current)${result.project ? ` — ${result.project}` : ""}`;
          const bodyText = changes.length ? formatProgrammeChanges(changes) : "No programme date changes found between the selected backup and current.";
          return { status: 200, jsonBody: { answer: `${title}\n\n${bodyText}` } };
        }

        return { status: 200, jsonBody: { answer: `${result.title}\n\n${result.body}` } };
      }

      // Free-text fallback: keep it deterministic and explicit.
      const help = [
        "This chat is fully deterministic (no AI).",
        "",
        "Use the Quick questions dropdown, or ask using this format:",
        "TEMPLATE:<id>",
        'PARAMS:{"project":"Holloway","dateFrom":"2026-01-01","dateTo":"2026-01-31"}',
        "",
        "Supported TEMPLATE ids:",
        "- overdue_first_issue_all / overdue_first_issue_project",
        "- overdue_statusA_all / overdue_statusA_project",
        "- statusA_approved_all / statusA_approved_project",
        "- done_all / done_project",
        "- programme_range_all / programme_range_project",
        "- compare_programme_all / compare_programme_project",
        "- comments_all / comments_project",
      ].join("\n");

      return { status: 200, jsonBody: { answer: help } };
    } catch (e) {
      return { status: 500, jsonBody: { error: "deterministic chat handler error", details: String(e) } };
    }
  },
});
