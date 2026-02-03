
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

/**
 * Deterministic /api/chat v2 (NO LLM)
 * Reads state from Azure Blob: container=workback, blob=state.json (defaults)
 * Adds follow-ups via queryState (stateless): client should send back queryState it receives.
 *
 * Response includes { queryState } but UI can ignore it without breaking.
 */

/* ------------------------------ Blob helpers ------------------------------ */

function getConnString() {
  return process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
}

function getContainerName() {
  return process.env.STATE_CONTAINER || "workback";
}

function getStateBlobName() {
  return process.env.STATE_BLOB || "state.json";
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

async function loadStateJson() {
  const conn = getConnString();
  if (!conn) throw new Error("Missing BLOB_CONNECTION_STRING or AzureWebJobsStorage");
  const service = BlobServiceClient.fromConnectionString(conn);
  const container = service.getContainerClient(getContainerName());
  const blob = container.getBlobClient(getStateBlobName());
  const dl = await blob.download();
  const text = await streamToString(dl.readableStreamBody);
  return text ? JSON.parse(text) : {};
}

/* ------------------------------ Date helpers ----------------------------- */

function safeStr(v) {
  return (v ?? "").toString().trim();
}

function isoToDate(iso) {
  if (!iso) return null;
  const parts = String(iso).split("-");
  if (parts.length < 3) return null;
  const y = Number(parts[0]), m = Number(parts[1]), d = Number(parts[2]);
  if (!y || !m || !d) return null;
  return new Date(Date.UTC(y, m - 1, d));
}

function addDaysISO(iso, days) {
  const dt = isoToDate(iso);
  if (!dt || typeof days !== "number" || Number.isNaN(days)) return "";
  return new Date(dt.getTime() + days * 86400000).toISOString().slice(0, 10);
}

function subDaysISO(iso, days) {
  return addDaysISO(iso, -days);
}

/* ------------------------- Flatten state into tables ---------------------- */

function buildIndexFromState(state) {
  const projects = Array.isArray(state?.projects) ? state.projects : [];
  const globalReqToStatusA = Number(state?.globalDaysReqToStatusA ?? 0);
  const globalStatusAToFirstIssue = Number(state?.globalDaysStatusAToFirstIssue ?? 0);

  const programme = [];
  const rows = [];
  const dict = { projects: new Set(), suppliers: new Set(), responsibilities: new Set() };

  projects.forEach((p) => {
    const projectName = safeStr(p?.name);
    if (projectName) dict.projects.add(projectName);

    // Programme/master
    (p.master || []).forEach((m) => {
      const blockZone = safeStr(m?.blockZone);
      (m.levels || []).forEach((lv) => {
        programme.push({
          project: projectName,
          blockZone,
          level: safeStr(lv?.name),
          startDate: safeStr(lv?.startDate),
          finishDate: safeStr(lv?.finishDate),
        });
      });
    });

    // Tracker rows
    const responsibilities = Array.isArray(p.responsibilities) ? p.responsibilities : [];
    const supplierByRespId = new Map(responsibilities.map((r) => [r.id, safeStr(r.supplier)]));

    responsibilities.forEach((r) => {
      const sup = safeStr(r?.supplier);
      const resp = safeStr(r?.name || r?.code);
      if (sup) dict.suppliers.add(sup);
      if (resp) dict.responsibilities.add(resp);
    });

    const pages = Array.isArray(p.pages) ? p.pages : [];
    pages
      .filter((pg) => !pg?.meta?.isMaster)
      .forEach((pg) => {
        const supplier = supplierByRespId.get(pg?.meta?.responsibilityId) || "—";
        if (supplier && supplier !== "—") dict.suppliers.add(supplier);

        const responsibility = safeStr(pg?.name);
        if (responsibility) dict.responsibilities.add(responsibility);

        (pg.rows || [])
          .filter((r) => r && r.kind !== "header")
          .forEach((r) => {
            if (r.notRequired) return;

            const requiredOnSite = safeStr(r.requiredOnSite || r.requiredOnSiteISO || r.anchorDateISO);

            const daysReqToStatusA = Number(r.overrideDaysReqToStatusA ?? globalReqToStatusA);
            const daysStatusAToFirstIssue = Number(r.overrideDaysStatusAToFirstIssue ?? globalStatusAToFirstIssue);

            const statusA =
              safeStr(r.statusA || r.statusAISO) ||
              (requiredOnSite ? subDaysISO(requiredOnSite, daysReqToStatusA) : "");

            const firstIssue =
              safeStr(r.firstIssue || r.firstIssueISO) ||
              (statusA ? subDaysISO(statusA, daysStatusAToFirstIssue) : "");

            rows.push({
              project: projectName,
              responsibility,
              supplier,
              item: safeStr(r.item),
              requiredOnSite,
              statusA,
              firstIssue,
              ticks: {
                statusA: !!r.statusADone,
                firstIssue: !!r.firstIssueDone,
                completed: !!r.completed,
                notRequired: !!r.notRequired,
              },
            });
          });
      });
  });

  return {
    programme,
    rows,
    dict: {
      projects: [...dict.projects].sort((a, b) => a.localeCompare(b)),
      suppliers: [...dict.suppliers].sort((a, b) => a.localeCompare(b)),
      responsibilities: [...dict.responsibilities].sort((a, b) => a.localeCompare(b)),
    },
  };
}

/* ------------------------------ Question parsing -------------------------- */

function getLastUserText(messages) {
  if (!Array.isArray(messages)) return "";
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i]?.role === "user") return String(messages[i]?.content || "");
  }
  return "";
}

function normalize(s) {
  return String(s || "").toLowerCase();
}

function findMention(text, candidates) {
  const t = normalize(text);
  const sorted = [...candidates].sort((a, b) => b.length - a.length);
  for (const c of sorted) {
    const cNorm = normalize(c);
    if (!cNorm) continue;
    if (cNorm.length >= 2 && t.includes(cNorm)) return c;
  }
  return "";
}

function parseLimit(text, fallback) {
  const t = String(text || "");
  const m = t.match(/\blimit\s*=\s*(\d+)\b/i) || t.match(/\btop\s+(\d+)\b/i);
  if (!m) return fallback;
  const n = Number(m[1]);
  return Number.isFinite(n) && n > 0 ? Math.min(n, 200) : fallback;
}

function monthKeyFromText(text, defaultYear) {
  const t = normalize(text);
  const months = {
    jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4,
    may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9,
    september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12,
  };
  const monthWord = Object.keys(months).find((k) => new RegExp(`\\b${k}\\b`).test(t));
  if (!monthWord) return null;

  const yearMatch = t.match(/\b(20\d{2})\b/);
  const year = yearMatch ? Number(yearMatch[1]) : Number(defaultYear);
  const mm = String(months[monthWord]).padStart(2, "0");
  return `${year}-${mm}`;
}

function startEndOfMonth(monthKey) {
  const [yS, mS] = String(monthKey).split("-");
  const y = Number(yS), m = Number(mS);
  const start = new Date(Date.UTC(y, m - 1, 1)).toISOString().slice(0, 10);
  const end = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
  return { start, end };
}

function overlapsMonth(startDate, finishDate, monthStart, monthEnd) {
  const s = isoToDate(startDate), f = isoToDate(finishDate);
  const ms = isoToDate(monthStart), me = isoToDate(monthEnd);
  if (!s || !f || !ms || !me) return false;
  return s <= me && f >= ms;
}

function isOverdue(row, todayISO) {
  if (row.ticks?.completed || row.ticks?.notRequired) return false;

  if (row.statusA && isoToDate(row.statusA) && !row.ticks?.statusA) {
    return isoToDate(row.statusA) < isoToDate(todayISO);
  }
  if (row.requiredOnSite && isoToDate(row.requiredOnSite)) {
    return isoToDate(row.requiredOnSite) < isoToDate(todayISO);
  }
  return false;
}

function isLikelyFollowUp(text) {
  const t = normalize(text);
  return /\b(only|just|what about|and|instead|filter|show me|those|them|in)\b/.test(t);
}

/**
 * queryState shape returned to client:
 * {
 *  v: 2,
 *  intent: string,
 *  dataset: "tracker"|"programme",
 *  filters: { project?: string, supplier?: string, overdue?: boolean },
 *  params: { monthKey?: string, limit?: number, dateField?: string }
 * }
 */

function parseNewQuestion(question, dict, uiContext) {
  const q = String(question || "").trim();
  const ql = normalize(q);

  const todayISO = safeStr(uiContext?.today) || new Date().toISOString().slice(0, 10);
  const activeProject = safeStr(uiContext?.activeProject?.name);

  const projectMention = findMention(q, dict.projects) || activeProject;
  const supplierMention = findMention(q, dict.suppliers);

  const limit = parseLimit(q, 25);

  // Programme
  if (/\b(on\s*site|onsite|programme|master)\b/.test(ql)) {
    const mk = monthKeyFromText(q, todayISO.slice(0, 4));
    if (mk) {
      return {
        v: 2,
        intent: "programme_month",
        dataset: "programme",
        filters: { project: projectMention || "" },
        params: { monthKey: mk, limit: 80 },
      };
    }
  }

  // Tracker intents
  if (ql.includes("supplier") && (ql.includes("most") || ql.includes("top")) && ql.includes("overdue")) {
    return {
      v: 2,
      intent: "group_count_supplier_overdue",
      dataset: "tracker",
      filters: { project: projectMention || "", overdue: true },
      params: { limit: 20 },
    };
  }

  if (ql.includes("overdue") && (ql.includes("next") || ql.includes("soonest") || ql.includes("upcoming"))) {
    return {
      v: 2,
      intent: "next_overdue",
      dataset: "tracker",
      filters: { project: projectMention || "", supplier: supplierMention || "", overdue: true },
      params: { dateField: "requiredOnSite", limit },
    };
  }

  if ((ql.includes("status a") || ql.includes("statusa")) && (ql.includes("next") || ql.includes("upcoming"))) {
    return {
      v: 2,
      intent: "next_statusA",
      dataset: "tracker",
      filters: { project: projectMention || "", supplier: supplierMention || "" },
      params: { dateField: "statusA", limit },
    };
  }

  if (ql.includes("how many") || ql.startsWith("count") || ql.includes("count")) {
    const overdue = ql.includes("overdue") ? true : undefined;
    return {
      v: 2,
      intent: "count_items",
      dataset: "tracker",
      filters: {
        project: projectMention || "",
        supplier: supplierMention || "",
        ...(overdue !== undefined ? { overdue } : {}),
      },
      params: {},
    };
  }

  return { v: 2, intent: "help", dataset: "tracker", filters: { project: projectMention || "" }, params: {} };
}

function applyFollowUp(prev, question, dict, uiContext) {
  const q = String(question || "").trim();
  const ql = normalize(q);

  const next = JSON.parse(JSON.stringify(prev || {}));
  next.v = 2;

  const activeProject = safeStr(uiContext?.activeProject?.name);
  const projectMention = findMention(q, dict.projects) || "";
  const supplierMention = findMention(q, dict.suppliers) || "";

  next.filters = { ...(next.filters || {}) };
  next.params = { ...(next.params || {}) };

  if (projectMention) next.filters.project = projectMention;
  else if (activeProject && !next.filters.project) next.filters.project = activeProject;

  if (supplierMention) next.filters.supplier = supplierMention;

  const lim = parseLimit(q, null);
  if (lim !== null) next.params.limit = lim;

  const todayISO = safeStr(uiContext?.today) || new Date().toISOString().slice(0, 10);
  const mk = monthKeyFromText(q, todayISO.slice(0, 4));
  if (mk && next.intent === "programme_month") next.params.monthKey = mk;

  // Explicit switches
  if (ql.includes("overdue")) next.filters.overdue = true;
  if (ql.includes("show") && (ql.includes("items") || ql.includes("those"))) {
    if (next.intent === "group_count_supplier_overdue") next.intent = "next_overdue";
  }

  return next;
}

/* ------------------------------ Execution -------------------------------- */

function execute(plan, index, uiContext) {
  const todayISO = safeStr(uiContext?.today) || new Date().toISOString().slice(0, 10);
  const filters = plan?.filters || {};
  const params = plan?.params || {};

  if (plan.intent === "programme_month") {
    const mk = params.monthKey;
    if (!mk) return { answer: "Please specify a month (e.g. February 2026).", data: { intent: plan.intent } };

    const { start, end } = startEndOfMonth(mk);
    const scoped = filters.project
      ? index.programme.filter((p) => normalize(p.project) === normalize(filters.project))
      : index.programme;

    const matches = scoped.filter((p) => overlapsMonth(p.startDate, p.finishDate, start, end));
    const outRows = matches.slice(0, params.limit || 80);
    const out = outRows.map((p) => `${p.project} — ${p.blockZone} — ${p.level} | ${p.startDate} → ${p.finishDate}`);

    return {
      answer: out.length
        ? `On site in ${mk}${filters.project ? ` for ${filters.project}` : ""}:
• ${out.join("\n• ")}`
        : `No programme items found on site in ${mk}${filters.project ? ` for ${filters.project}` : ""}.`,
      data: { intent: plan.intent, monthKey: mk, count: matches.length, rows: outRows },
    };
  }

  // Tracker scope
  let rows = index.rows;
  if (filters.project) rows = rows.filter((r) => normalize(r.project) === normalize(filters.project));
  if (filters.supplier) rows = rows.filter((r) => normalize(r.supplier) === normalize(filters.supplier));

  if (plan.intent === "group_count_supplier_overdue") {
    const overdue = rows.filter((r) => isOverdue(r, todayISO));
    const counts = new Map();
    overdue.forEach((r) => counts.set(r.supplier, (counts.get(r.supplier) || 0) + 1));
    const ranked = [...counts.entries()].sort((a, b) => b[1] - a[1]).slice(0, params.limit || 20);

    if (!ranked.length) return { answer: `No overdue items found${filters.project ? ` in ${filters.project}` : ""}.`, data: { intent: plan.intent } };

    const [topSupplier, topCount] = ranked[0];
    return {
      answer:
        `Supplier with the most overdue items${filters.project ? ` in ${filters.project}` : ""}: ${topSupplier} (${topCount}).\n` +
        `Top suppliers:\n• ${ranked.map(([s, c]) => `${s}: ${c}`).join("\n• ")}`,
      data: { intent: plan.intent, ranked: ranked.map(([supplier, count]) => ({ supplier, count })) },
    };
  }

  if (plan.intent === "next_overdue") {
    let scoped = rows;
    if (filters.overdue) scoped = scoped.filter((r) => isOverdue(r, todayISO));
    scoped.sort((a, b) => (a.requiredOnSite || "").localeCompare(b.requiredOnSite || ""));
    const outRows = scoped.slice(0, params.limit || 25);
    const out = outRows.map((r) => `${r.project} — ${r.supplier} — ${r.item} | required on site: ${r.requiredOnSite} | status A: ${r.statusA}`);

    return {
      answer: out.length
        ? `Overdue next${filters.project ? ` in ${filters.project}` : ""}${filters.supplier ? ` (${filters.supplier})` : ""}:\n• ${out.join("\n• ")}`
        : `No overdue items found${filters.project ? ` in ${filters.project}` : ""}.`,
      data: { intent: plan.intent, count: scoped.length, rows: outRows },
    };
  }

  if (plan.intent === "next_statusA") {
    const pending = rows.filter((r) => r.statusA && !r.ticks?.statusA);
    pending.sort((a, b) => (a.statusA || "").localeCompare(b.statusA || ""));
    const outRows = pending.slice(0, params.limit || 25);
    const out = outRows.map((r) => `${r.statusA} — ${r.supplier} — ${r.item}`);

    return {
      answer: out.length
        ? `Next Status A dates${filters.project ? ` in ${filters.project}` : ""}${filters.supplier ? ` (${filters.supplier})` : ""}:\n• ${out.join("\n• ")}`
        : `No upcoming Status A dates found${filters.project ? ` in ${filters.project}` : ""}.`,
      data: { intent: plan.intent, count: pending.length, rows: outRows },
    };
  }

  if (plan.intent === "count_items") {
    let scoped = rows;
    if (filters.overdue === true) scoped = scoped.filter((r) => isOverdue(r, todayISO));
    const n = scoped.length;

    return {
      answer:
        `Count${filters.overdue ? " (overdue)" : ""}` +
        `${filters.project ? ` in ${filters.project}` : ""}` +
        `${filters.supplier ? ` for ${filters.supplier}` : ""}: ${n}.`,
      data: { intent: plan.intent, count: n },
    };
  }

  return {
    answer:
      `I can answer programme + tracker questions deterministically.\n` +
      `Try:\n` +
      `• What is overdue next?\n` +
      `• Which supplier has the most overdue items?\n` +
      `• What are the next Status A dates?\n` +
      `• Count overdue items\n` +
      `• What is on site in February 2026?\n\n` +
      `Follow-ups you can ask:\n` +
      `• Only in Holloway\n` +
      `• Just Keyfix\n` +
      `• Top 10\n` +
      `• What about March 2026`,
    data: { intent: "help" },
  };
}

/* ------------------------------ Azure Function ---------------------------- */

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const body = await req.json().catch(() => ({}));
      const question = getLastUserText(body?.messages);
      const uiContext = body?.context || {};

      const state = await loadStateJson();
      const index = buildIndexFromState(state);

      const incoming = body?.queryState;
      const plan =
        incoming && isLikelyFollowUp(question)
          ? applyFollowUp(incoming, question, index.dict, uiContext)
          : parseNewQuestion(question, index.dict, uiContext);

      const result = execute(plan, index, uiContext);

      return {
        status: 200,
        jsonBody: {
          answer: result.answer,
          data: result.data,
          queryState: plan,
        },
      };
    } catch (err) {
      return { status: 500, jsonBody: { error: err?.message || String(err) } };
    }
  },
});
