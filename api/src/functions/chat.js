function YESNO(b){return b?"YES":"NO";}
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

/* ---------------- context compaction ---------------- */

function truncateStr(s, max = 400) {
  const v = (s ?? "").toString();
  if (v.length <= max) return v;
  return v.slice(0, max) + "…";
}

function jsonByteLen(x) {
  try {
    return Buffer.byteLength(JSON.stringify(x));
  } catch {
    return Infinity;
  }
}

function compactRows(rows, { includeComments, searchTextMax }) {
  const out = [];
  for (const r of Array.isArray(rows) ? rows : []) {
    out.push({
      key: safeStr(r?.key),
      project: safeStr(r?.project),
      responsibility: safeStr(r?.responsibility),
      supplier: safeStr(r?.supplier),
      item: safeStr(r?.item),
      requiredOnSite: safeStr(r?.requiredOnSite),
      statusA: safeStr(r?.statusA),
      firstIssue: safeStr(r?.firstIssue),
      timeframe: safeStr(r?.timeframe),
      totalDurationDays: r?.totalDurationDays ?? null,
      ticks: r?.ticks && typeof r.ticks === "object" ? r.ticks : (r?.ticksBool ?? null),
      comments: includeComments ? truncateStr(r?.comments, 500) : undefined,
      commentsCount: r?.commentsCount ?? undefined,
      status: safeStr(r?.status),
      traffic: safeStr(r?.traffic),
      searchText: truncateStr(r?.searchText, searchTextMax),
    });
  }
  return out;
}

function compactContext(appContext) {
  // Keep behaviour, but avoid OpenAI request failures caused by extremely large context.
  const ctx = (appContext && typeof appContext === "object") ? appContext : {};
  const base = {
    today: safeStr(ctx.today),
    programme: Array.isArray(ctx.programme) ? ctx.programme : [],
    programmeIndex: (ctx.programmeIndex && typeof ctx.programmeIndex === "object") ? ctx.programmeIndex : {},
    rows: Array.isArray(ctx.rows) ? ctx.rows : [],
  };

  // Attempt progressively more compact representations until the payload is a sane size.
  const MAX_BYTES = Number(process.env.OPENAI_CONTEXT_MAX_BYTES || 350_000);

  // 1) full context as-is (if small)
  if (jsonByteLen(base) <= MAX_BYTES) return base;

  // 2) compact rows (keep comments)
  const v2 = {
    today: base.today,
    programme: base.programme,
    programmeIndex: base.programmeIndex,
    rows: compactRows(base.rows, { includeComments: true, searchTextMax: 500 }),
  };
  if (jsonByteLen(v2) <= MAX_BYTES) return v2;

  // 3) compact rows (drop comments)
  const v3 = {
    today: base.today,
    programme: base.programme,
    programmeIndex: base.programmeIndex,
    rows: compactRows(base.rows, { includeComments: false, searchTextMax: 450 }),
  };
  if (jsonByteLen(v3) <= MAX_BYTES) return v3;

  // 4) keep only authoritative indices + minimal row fields
  const v4 = {
    today: base.today,
    programmeIndex: base.programmeIndex,
    rows: compactRows(base.rows, { includeComments: false, searchTextMax: 320 }).map((r) => ({
      key: r.key,
      project: r.project,
      responsibility: r.responsibility,
      supplier: r.supplier,
      item: r.item,
      requiredOnSite: r.requiredOnSite,
      statusA: r.statusA,
      firstIssue: r.firstIssue,
      timeframe: r.timeframe,
      ticks: r.ticks,
      status: r.status,
      traffic: r.traffic,
      searchText: r.searchText,
    })),
  };
  return v4;
}

/* ---------------- blob helpers ---------------- */

function getBlobClients() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";

  if (!conn) throw new Error("Storage connection string is not set (BLOB_CONNECTION_STRING or AzureWebJobsStorage)");

  const service = BlobServiceClient.fromConnectionString(conn);
  const container = service.getContainerClient(containerName);
  return { container };
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

function extractBackupDateFromName(name) {
  // backups/state-YYYY-MM-DD.json
  const m = String(name || "").match(/state-(\d{4}-\d{2}-\d{2})\.json$/);
  return m ? m[1] : null;
}

function safeStr(v) {
  return (v ?? "").toString().trim();
}

function isoToUtcDate(s) {
  if (!s) return null;
  const parts = String(s).split("-").map(Number);
  if (parts.length !== 3) return null;
  const [y, m, d] = parts;
  if (!y || !m || !d) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function diffDaysUTC(aISO, bISO) {
  const a = isoToUtcDate(aISO);
  const b = isoToUtcDate(bISO);
  if (!a || !b) return null;
  return Math.round((b.getTime() - a.getTime()) / 86400000);
}

const MONTHS = {
  jan: 1, january: 1,
  feb: 2, february: 2,
  mar: 3, march: 3,
  apr: 4, april: 4,
  may: 5,
  jun: 6, june: 6,
  jul: 7, july: 7,
  aug: 8, august: 8,
  sep: 9, sept: 9, september: 9,
  oct: 10, october: 10,
  nov: 11, november: 11,
  dec: 12, december: 12,
};

function monthKeyFromText(text, defaultYear) {
  const t = String(text || "").toLowerCase();

  // match month words
  const monthWord = Object.keys(MONTHS).find((k) => new RegExp(`\\b${k}\\b`).test(t));
  if (!monthWord) return null;

  // match year (e.g. 2026). If not present, use defaultYear.
  const yearMatch = t.match(/\b(20\d{2})\b/);
  const year = yearMatch ? Number(yearMatch[1]) : Number(defaultYear);

  const m = MONTHS[monthWord];
  const mm = String(m).padStart(2, "0");
  return `${year}-${mm}`; // e.g. 2026-02
}

function isProgrammeMonthQuery(text) {
  const t = String(text || "").toLowerCase();
  // “on site”, “onsite”, “programme”, “stage” etc.
  return /\bon\s*site\b|\bonsite\b|\bprogramme\b|\bstage\b|\blevel\b/.test(t);
}

/* ---------------- index builders ---------------- */

// Builds a minimal “chat-like context” from a saved state.json payload.
function buildIndexFromState(state) {
  const projects = Array.isArray(state?.projects) ? state.projects : [];

  // programme: master levels start/finish/duration
  const programme = projects.flatMap((p) =>
    (p.master || []).flatMap((m) =>
      (m.levels || []).map((lv) => ({
        project: safeStr(p.name),
        blockZone: safeStr(m.blockZone),
        level: safeStr(lv.name),
        startDate: safeStr(lv.startDate),
        finishDate: safeStr(lv.finishDate),
        durationDays: lv.startDate && lv.finishDate ? diffDaysUTC(lv.startDate, lv.finishDate) : null,
      }))
    )
  );

  // rows: tracker rows (exclude headers + notRequired like your summaryItems)
  const rows = [];
  projects.forEach((p) => {
    const responsibilities = Array.isArray(p.responsibilities) ? p.responsibilities : [];
    const pages = Array.isArray(p.pages) ? p.pages : [];

    const supplierByRespId = new Map(responsibilities.map((r) => [r.id, safeStr(r.supplier)]));

    pages
      .filter((pg) => !pg?.meta?.isMaster)
      .forEach((pg) => {
        (pg.rows || [])
          .filter((r) => r && r.kind !== "header")
          .filter((r) => !r.notRequired)
          .forEach((r) => {
            // These are “raw” row fields; your UI computes derived dates elsewhere,
            // so we store what the state likely has + ticks + comments.
            const comments = Array.isArray(r.comments) ? r.comments : [];
            const commentsText = comments
              .map((c) => [safeStr(c?.dateISO), safeStr(c?.name || c?.lockedBy), safeStr(c?.text)].filter(Boolean).join(" — "))
              .filter(Boolean)
              .join("\n");

            const supplier = supplierByRespId.get(pg?.meta?.responsibilityId) || "";

            // Create a best-effort key. Ideally you add IDs in chatContext later,
            // but this works well enough for comparisons.
            const key = [safeStr(p.name), safeStr(pg.name), safeStr(supplier), safeStr(r.item)].join(" | ");

            rows.push({
              key,
              project: safeStr(p.name),
              responsibility: safeStr(pg.name),
              supplier: supplier || "—",
              item: safeStr(r.item),

              // Best effort date fields (if your state stores computed dates, include them)
              requiredOnSite: safeStr(r.requiredOnSite || r.requiredOnSiteISO),
              statusA: safeStr(r.statusA || r.statusAISO),
              firstIssue: safeStr(r.firstIssue || r.firstIssueISO),

              // If you store these in state (many builds do)
              timeframe: safeStr(r.timeframe),
              daysReqToStatusA: r.overrideDaysReqToStatusA ?? null,
              daysStatusAToFirstIssue: r.overrideDaysStatusAToFirstIssue ?? null,

              // ticks
              ticks: {
                statusA: !!r.statusADone,
                firstIssue: !!r.firstIssueDone,
                completed: !!r.completed,
                notRequired: !!r.notRequired,
              },

              comments: commentsText,
              commentsCount: comments.length,
            });
          });
      });
  });

  return { programme, rows };
}

// Diff current vs backup (tick + date changes)
function diffCurrentVsBackup(currentCtx, backupIndex) {
  const currentRows = Array.isArray(currentCtx?.rows) ? currentCtx.rows : [];
  const backupRows = Array.isArray(backupIndex?.rows) ? backupIndex.rows : [];

  const curMap = new Map();
  currentRows.forEach((r) => {
    const key =
      safeStr(r.key) ||
      [safeStr(r.project), safeStr(r.responsibility), safeStr(r.supplier), safeStr(r.item)].join(" | ");
    curMap.set(key, r);
  });

  const changes = [];
  backupRows.forEach((b) => {
    const key = safeStr(b.key) || [safeStr(b.project), safeStr(b.responsibility), safeStr(b.supplier), safeStr(b.item)].join(" | ");
    const c = curMap.get(key);
    if (!c) return;

    const tickChanges = [];
    const ct = c.ticks || {};
    const bt = b.ticks || {};

    ["statusA", "firstIssue", "completed", "notRequired"].forEach((k) => {
      const cv = !!ct[k];
      const bv = !!bt[k];
      if (cv !== bv) tickChanges.push({ tick: k, from: bv, to: cv });
    });

    const dateChanges = [];
    ["requiredOnSite", "statusA", "firstIssue", "timeframe"].forEach((k) => {
      const cv = safeStr(c[k]);
      const bv = safeStr(b[k]);
      if (cv !== bv && (cv || bv)) dateChanges.push({ field: k, from: bv || null, to: cv || null });
    });

    if (tickChanges.length || dateChanges.length) {
      changes.push({
        key,
        project: safeStr(c.project),
        responsibility: safeStr(c.responsibility),
        supplier: safeStr(c.supplier),
        item: safeStr(c.item),
        tickChanges,
        dateChanges,
      });
    }
  });

  // Keep payload small
  return {
    changedCount: changes.length,
    changedSample: changes.slice(0, 120),
  };
}

/* ---------------- chat function ---------------- */

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const body = await req.json().catch(() => ({}));
      const messages = body?.messages;
      const appContextRaw = body?.context;
      const searchBackups = !!body?.searchBackups;

      // The UI can send a very large context object. To prevent intermittent
      // OpenAI 4xx errors (which surface as 500 to the UI), we compact the
      // context server-side while preserving the important fields.
      const appContext = compactContext(appContextRaw);

      const apiKey = process.env.OPENAI_API_KEY;
      const model = process.env.OPENAI_MODEL || "gpt-4o-mini";

      if (!apiKey) {
        return {
          status: 500,
          jsonBody: {
            error:
              "OPENAI_API_KEY is not set. Add it in Azure Static Web Apps > Configuration > Application settings.",
          },
        };
      }

      // Optional: load + index backups (only when user ticks "search backups")
      const backupsPayload = { enabled: searchBackups, count: 0, snapshots: [], comparisons: [] };

      if (searchBackups) {
        const prefix = process.env.BACKUP_PREFIX || "backups/";
        const maxBackups = Number(process.env.BACKUP_CHAT_LIMIT || 6);

        const { container } = getBlobClients();
        await container.createIfNotExists();

        const found = [];
        for await (const blob of container.listBlobsFlat({ prefix })) {
          const date = extractBackupDateFromName(blob.name);
          if (!date) continue;
          found.push({ name: blob.name, date });
        }

        found.sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : 0));

        const selected = found.slice(0, Math.max(0, maxBackups));
        const snapshots = [];

        for (const b of selected) {
          try {
            const client = container.getBlobClient(b.name);
            const dl = await client.download();
            const text = await streamToString(dl.readableStreamBody);
            const state = text ? JSON.parse(text) : null;

            const index = buildIndexFromState(state || {});
            snapshots.push({
              date: b.date,
              blobName: b.name,
              programme: index.programme,
              rows: index.rows,
            });
          } catch {
            // ignore bad blobs
          }
        }

        backupsPayload.count = snapshots.length;
        backupsPayload.snapshots = snapshots.map((s) => ({
          date: s.date,
          blobName: s.blobName,
          // keep the data, but you can trim if needed
          programme: s.programme,
          rows: s.rows,
        }));

        // Precompute diffs current vs each backup (helps model answer comparisons fast)
        backupsPayload.comparisons = snapshots.map((s) => ({
          date: s.date,
          blobName: s.blobName,
          diff: diffCurrentVsBackup(appContext || {}, s),
        }));
      }

      const systemText = `
You are an expert assistant for the SupplySync programme tracker.

STRICT RULES:
- Always search ALL rows in APP_CONTEXT_JSON.rows unless the user explicitly limits scope.
- Treat rows[].searchText as the primary full-text index.
- Do NOT answer from sample data if rows[] exists.
- Tick boxes are authoritative in rows[].ticks (YES/NO).
- Dates may appear as requiredOnSite, statusA, firstIssue, timeframe, or duration fields.
- When asked "is X ticked", respond YES or NO and name the item.
- When asked to list items, return ALL matching rows or ask the user to narrow scope.
- Never guess. If data is missing, say so clearly.

PROGRAMME RULES (master schedule):
- Programme questions must use APP_CONTEXT_JSON.programme and APP_CONTEXT_JSON.programmeIndex (not rows).
- “Start” means programme[].startDate. “Finish” means programme[].finishDate.
- “On site” means ANY overlap between startDate and finishDate (inclusive).
- For “on site in <month>” questions:
  - ALWAYS use APP_CONTEXT_JSON.programmeIndex[YYYY-MM]
  - programmeIndex is COMPLETE and AUTHORITATIVE
  - Do NOT infer from programme[] when programmeIndex exists
  - If stages are requested, use programmeIndex[YYYY-MM].stages.
  - If PROGRAMME_INDEX_MATCH is present, it is complete and authoritative: use it first and do not omit items.

BACKUPS:
- If BACKUPS_JSON is present, you may compare snapshots.
- Always state the backup date(s) explicitly (YYYY-MM-DD).
- When comparing, show BEFORE → AFTER values and list only changed items unless asked for full dumps.

Prefer accuracy over brevity.
`;


      
      // Deterministic programme month slice for “on site in <month>” questions (prevents omissions)
      const userText =
        Array.isArray(messages) && messages.length
          ? String(messages[messages.length - 1]?.content || "")
          : "";

      const todayISO = safeStr(appContext?.today) || new Date().toISOString().slice(0, 10);
      const fallbackYear = todayISO.slice(0, 4);

      let programmeIndexMatch = null;
      if (isProgrammeMonthQuery(userText)) {
        const mk = monthKeyFromText(userText, fallbackYear);
        const idx = appContext?.programmeIndex || null;

        if (mk && idx && idx[mk]) {
          programmeIndexMatch = {
            monthKey: mk,
            projects: idx[mk]?.projects || [],
            stages: idx[mk]?.stages || [],
            projectCount: Array.isArray(idx[mk]?.projects) ? idx[mk].projects.length : 0,
            stageCount: Array.isArray(idx[mk]?.stages) ? idx[mk].stages.length : 0,
            instruction:
              "Use PROGRAMME_INDEX_MATCH as complete authoritative data for 'on site in this month' questions. Do not omit any projects/stages.",
          };
        }
      }

      const input = [
        { role: "system", content: systemText },
        { role: "system", content: `APP_CONTEXT_JSON:\n${JSON.stringify(appContext ?? {}, null, 2)}` },
        ...(programmeIndexMatch
          ? [{ role: "system", content: `PROGRAMME_INDEX_MATCH:\n${JSON.stringify(programmeIndexMatch, null, 2)}` }]
          : []),
        ...(searchBackups
          ? [{ role: "system", content: `BACKUPS_JSON:\n${JSON.stringify(backupsPayload ?? {}, null, 2)}` }]
          : []),
        ...(Array.isArray(messages) ? messages : []),
      ];

      const r = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model,
          input,
          temperature: 0.2,
          max_output_tokens: 900,
        }),
      });

      const data = await r.json().catch(() => ({}));

      if (!r.ok) {
        return {
          status: 500,
          jsonBody: {
            error: "OpenAI request failed",
            status: r.status,
            details: data,
          },
        };
      }

      const answer =
        data?.output_text ||
        data?.output?.[0]?.content?.[0]?.text ||
        "No response text returned.";

      return {
        status: 200,
        jsonBody: { answer },
      };
    } catch (e) {
      return {
        status: 500,
        jsonBody: { error: "chat handler error", details: String(e) },
      };
    }
  },
});

