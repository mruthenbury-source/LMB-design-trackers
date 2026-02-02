import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

/**
 * Deterministic chat endpoint (no LLM).
 * Reads authoritative state from Azure Blob Storage (container: workback, blob: state.json),
 * flattens into queryable "tables", parses the user's question using rules, and returns
 * deterministic answers + optional structured data.
 *
 * Env vars:
 * - BLOB_CONNECTION_STRING (preferred) or AzureWebJobsStorage
 * - STATE_CONTAINER (optional, default "workback")
 * - STATE_BLOB (optional, default "state.json")
 */

/* ---------------- blob helpers ---------------- */

function getContainerClient() {
  const conn =
    process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  if (!conn) {
    throw new Error(
      "Missing BLOB_CONNECTION_STRING or AzureWebJobsStorage for blob access."
    );
  }
  const service = BlobServiceClient.fromConnectionString(conn);
  return service.getContainerClient(containerName);
}

async function streamToString(readable) {
  if (!readable) return "";
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

async function loadStateJson() {
  const container = getContainerClient();
  const blobName = process.env.STATE_BLOB || "state.json";
  const blob = container.getBlobClient(blobName);

  const dl = await blob.download();
  const text = await streamToString(dl.readableStreamBody);
  return text ? JSON.parse(text) : {};
}

/* ---------------- deterministic indexing ---------------- */

function safeStr(v) {
  return (v ?? "").toString().trim();
}

function isoToDate(s) {
  if (!s) return null;
  const parts = String(s).split("-");
  if (parts.length !== 3) return null;
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);
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

function buildIndexFromState(state) {
  const projects = Array.isArray(state?.projects) ? state.projects : [];

  const globalReqToStatusA = Number(state?.globalDaysReqToStatusA ?? 0);
  const globalStatusAToFirstIssue = Number(
    state?.globalDaysStatusAToFirstIssue ?? 0
  );

  // Programme schedule rows (master)
  const programme = [];
  // Tracker rows (pages)
  const rows = [];

  for (const p of projects) {
    const projectName = safeStr(p?.name);

    // Programme: project.master[].levels[]
    const masterBlocks = Array.isArray(p?.master) ? p.master : [];
    for (const m of masterBlocks) {
      const blockZone = safeStr(m?.blockZone);
      const levels = Array.isArray(m?.levels) ? m.levels : [];
      for (const lv of levels) {
        programme.push({
          project: projectName,
          blockZone,
          level: safeStr(lv?.name),
          startDate: safeStr(lv?.startDate),
          finishDate: safeStr(lv?.finishDate),
        });
      }
    }

    // Tracker: map responsibilityId -> supplier
    const responsibilities = Array.isArray(p?.responsibilities)
      ? p.responsibilities
      : [];
    const supplierByRespId = new Map(
      responsibilities.map((r) => [r?.id, safeStr(r?.supplier)])
    );

    const pages = Array.isArray(p?.pages) ? p.pages : [];
    for (const pg of pages) {
      // Skip master pages (programme schedule lives in p.master)
      if (pg?.meta?.isMaster) continue;

      const responsibility = safeStr(pg?.name);
      const supplier =
        supplierByRespId.get(pg?.meta?.responsibilityId) || "—";

      const pageRows = Array.isArray(pg?.rows) ? pg.rows : [];
      for (const r of pageRows) {
        if (!r || r.kind === "header") continue;
        if (r.notRequired) continue;

        // Required on site date: prefer explicit fields, fallback to anchorDateISO
        const requiredOnSite = safeStr(
          r.requiredOnSite || r.requiredOnSiteISO || r.anchorDateISO
        );

        const daysReqToStatusA = Number(
          r.overrideDaysReqToStatusA ?? globalReqToStatusA
        );
        const daysStatusAToFirstIssue = Number(
          r.overrideDaysStatusAToFirstIssue ?? globalStatusAToFirstIssue
        );

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
      }
    }
  }

  // Dictionaries for parsing
  const dict = {
    projects: [...new Set(rows.map((r) => r.project).filter(Boolean))].sort(),
    suppliers: [...new Set(rows.map((r) => r.supplier).filter(Boolean))].sort(),
    responsibilities: [
      ...new Set(rows.map((r) => r.responsibility).filter(Boolean)),
    ].sort(),
  };

  return { programme, rows, dict };
}

/* ---------------- deterministic parsing + execution ---------------- */

function getLastUserText(messages) {
  if (!Array.isArray(messages)) return "";
  for (let i = messages.length - 1; i >= 0; i--) {
    if (messages[i]?.role === "user") return String(messages[i]?.content || "");
  }
  return "";
}

function normalize(s) {
  return safeStr(s).toLowerCase();
}

function pickBestMatch(text, candidates) {
  const t = normalize(text);
  // exact (case-insensitive) containment match
  let best = "";
  let bestLen = 0;
  for (const c of candidates) {
    const cn = normalize(c);
    if (!cn) continue;
    // whole-word-ish match
    const re = new RegExp(`\\b${escapeRegExp(cn)}\\b`, "i");
    if (re.test(t) && cn.length > bestLen) {
      best = c;
      bestLen = cn.length;
    }
  }
  return best || "";
}

function escapeRegExp(s) {
  return String(s).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function todayISOFromContext(ctx) {
  const iso = safeStr(ctx?.today);
  if (iso && isoToDate(iso)) return iso;
  // UTC date is fine for deterministic comparisons; UI already sends "today" typically.
  return new Date().toISOString().slice(0, 10);
}

function isOverdue(row, todayISO) {
  if (row?.ticks?.completed || row?.ticks?.notRequired) return false;

  // Prefer statusA if present and not done
  if (row.statusA && isoToDate(row.statusA) && !row?.ticks?.statusA) {
    return isoToDate(row.statusA) < isoToDate(todayISO);
  }
  // Fallback: requiredOnSite
  if (row.requiredOnSite && isoToDate(row.requiredOnSite)) {
    return isoToDate(row.requiredOnSite) < isoToDate(todayISO);
  }
  return false;
}

function monthKeyFromText(text, defaultYear) {
  const t = normalize(text);
  const months = {
    jan: 1,
    january: 1,
    feb: 2,
    february: 2,
    mar: 3,
    march: 3,
    apr: 4,
    april: 4,
    may: 5,
    jun: 6,
    june: 6,
    jul: 7,
    july: 7,
    aug: 8,
    august: 8,
    sep: 9,
    sept: 9,
    september: 9,
    oct: 10,
    october: 10,
    nov: 11,
    november: 11,
    dec: 12,
    december: 12,
  };
  let monthNum = null;
  for (const k of Object.keys(months)) {
    const re = new RegExp(`\\b${k}\\b`, "i");
    if (re.test(t)) {
      monthNum = months[k];
      break;
    }
  }
  if (!monthNum) return null;

  const yearMatch = t.match(/\b(20\d{2})\b/);
  const year = yearMatch ? Number(yearMatch[1]) : Number(defaultYear);
  const mm = String(monthNum).padStart(2, "0");
  return `${year}-${mm}`;
}

function startEndOfMonth(monthKey) {
  const [yS, mS] = String(monthKey).split("-");
  const y = Number(yS);
  const m = Number(mS);
  const start = new Date(Date.UTC(y, m - 1, 1)).toISOString().slice(0, 10);
  const end = new Date(Date.UTC(y, m, 0)).toISOString().slice(0, 10);
  return { start, end };
}

function overlapsMonth(startDate, finishDate, monthStart, monthEnd) {
  const s = isoToDate(startDate);
  const f = isoToDate(finishDate);
  const ms = isoToDate(monthStart);
  const me = isoToDate(monthEnd);
  if (!s || !f || !ms || !me) return false;
  return s <= me && f >= ms;
}

function formatListLines(lines) {
  if (!lines.length) return "";
  return `• ${lines.join("\n• ")}`;
}

function handleDeterministic(question, index, uiContext) {
  const q = safeStr(question);
  const ql = q.toLowerCase();
  const todayISO = todayISOFromContext(uiContext);

  // Default project from context if present
  const defaultProject = safeStr(uiContext?.activeProject?.name);
  const projectFromText = pickBestMatch(q, index.dict.projects);
  const project = projectFromText || defaultProject;

  // Supplier filter from text if present
  const supplierFromText = pickBestMatch(q, index.dict.suppliers);
  const supplier = supplierFromText;

  // --- Programme month queries ---
  if (/\b(on\s*site|onsite|programme)\b/i.test(q)) {
    const mk = monthKeyFromText(q, todayISO.slice(0, 4));
    if (mk) {
      const { start, end } = startEndOfMonth(mk);
      const scoped = project
        ? index.programme.filter(
            (p) => normalize(p.project) === normalize(project)
          )
        : index.programme;

      const matched = scoped.filter((p) =>
        overlapsMonth(p.startDate, p.finishDate, start, end)
      );

      const lines = matched.slice(0, 80).map(
        (p) =>
          `${p.project} — ${p.blockZone} — ${p.level} | ${p.startDate} → ${p.finishDate}`
      );

      return {
        answer: matched.length
          ? `On site in ${mk}${project ? ` for ${project}` : ""}:\n${formatListLines(lines)}`
          : `No programme items found on site in ${mk}${project ? ` for ${project}` : ""}.`,
        data: {
          intent: "programme_month",
          monthKey: mk,
          project: project || null,
          count: matched.length,
          rows: matched.slice(0, 80),
        },
      };
    }
  }

  // Prepare tracker scope
  let tracker = index.rows;

  if (project) {
    tracker = tracker.filter((r) => normalize(r.project) === normalize(project));
  }
  if (supplier) {
    tracker = tracker.filter(
      (r) => normalize(r.supplier) === normalize(supplier)
    );
  }

  // --- Supplier has most overdue ---
  if (
    ql.includes("supplier") &&
    (ql.includes("most") || ql.includes("top")) &&
    ql.includes("overdue")
  ) {
    const overdue = tracker.filter((r) => isOverdue(r, todayISO));
    const counts = new Map();
    for (const r of overdue) {
      counts.set(r.supplier, (counts.get(r.supplier) || 0) + 1);
    }
    const ranked = [...counts.entries()]
      .sort((a, b) => b[1] - a[1])
      .slice(0, 20);

    if (!ranked.length) {
      return {
        answer: `No overdue items found${project ? ` in ${project}` : ""}.`,
        data: { intent: "group_count_supplier_overdue", project: project || null },
      };
    }

    const [topSupplier, topCount] = ranked[0];
    const lines = ranked.map(([s, c]) => `${s}: ${c}`);
    return {
      answer:
        `Supplier with the most overdue items${project ? ` in ${project}` : ""}: ${topSupplier} (${topCount}).\n` +
        `Top suppliers:\n${formatListLines(lines)}`,
      data: {
        intent: "group_count_supplier_overdue",
        project: project || null,
        ranked: ranked.map(([s, c]) => ({ supplier: s, count: c })),
      },
    };
  }

  // --- Overdue next ---
  if (
    ql.includes("overdue") &&
    (ql.includes("next") || ql.includes("soonest") || ql.includes("upcoming"))
  ) {
    const overdue = tracker.filter((r) => isOverdue(r, todayISO));
    overdue.sort((a, b) =>
      safeStr(a.requiredOnSite).localeCompare(safeStr(b.requiredOnSite))
    );

    const out = overdue.slice(0, 25).map(
      (r) =>
        `${r.project} — ${r.supplier} — ${r.item} | required on site: ${r.requiredOnSite} | status A: ${r.statusA}`
    );

    return {
      answer: out.length
        ? `Overdue next${project ? ` in ${project}` : ""}:\n${formatListLines(out)}`
        : `No overdue items found${project ? ` in ${project}` : ""}.`,
      data: {
        intent: "next_overdue",
        project: project || null,
        supplier: supplier || null,
        totalOverdue: overdue.length,
        rows: overdue.slice(0, 25),
      },
    };
  }

  // --- Next Status A dates ---
  if (
    (ql.includes("status a") || ql.includes("statusa")) &&
    (ql.includes("next") || ql.includes("upcoming"))
  ) {
    const pending = tracker.filter((r) => r.statusA && !r?.ticks?.statusA);
    pending.sort((a, b) => safeStr(a.statusA).localeCompare(safeStr(b.statusA)));

    const out = pending.slice(0, 25).map(
      (r) => `${r.statusA} — ${r.supplier} — ${r.item}`
    );

    return {
      answer: out.length
        ? `Next Status A dates${project ? ` in ${project}` : ""}:\n${formatListLines(out)}`
        : `No upcoming Status A dates found${project ? ` in ${project}` : ""}.`,
      data: {
        intent: "next_statusA",
        project: project || null,
        supplier: supplier || null,
        count: pending.length,
        rows: pending.slice(0, 25),
      },
    };
  }

  // --- Count overdue (explicit or implied) ---
  if (ql.includes("how many") || ql.includes("count")) {
    const wantOverdue = ql.includes("overdue");
    if (wantOverdue) {
      const overdue = tracker.filter((r) => isOverdue(r, todayISO));
      return {
        answer: `Overdue count${project ? ` in ${project}` : ""}${
          supplier ? ` for ${supplier}` : ""
        }: ${overdue.length}.`,
        data: {
          intent: "count_overdue",
          project: project || null,
          supplier: supplier || null,
          count: overdue.length,
        },
      };
    }
    // total count
    return {
      answer: `Row count${project ? ` in ${project}` : ""}${
        supplier ? ` for ${supplier}` : ""
      }: ${tracker.length}.`,
      data: {
        intent: "count_rows",
        project: project || null,
        supplier: supplier || null,
        count: tracker.length,
      },
    };
  }

  // Deterministic help fallback
  return {
    answer:
      `I can answer programme + tracker questions deterministically.\n` +
      `Try:\n` +
      `• What is overdue next?\n` +
      `• Which supplier has the most overdue items?\n` +
      `• What are the next Status A dates?\n` +
      `• How many overdue items are there?\n` +
      `• What is on site in February 2026?`,
    data: { intent: "help" },
  };
}

/* ---------------- Azure Function ---------------- */

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const body = await req.json().catch(() => ({}));
      const question = getLastUserText(body?.messages);

      // Always read from blob (authoritative source) to avoid relying on huge UI context.
      const state = await loadStateJson();
      const index = buildIndexFromState(state);

      const result = handleDeterministic(question, index, body?.context || {});
      // Keep UI compatibility: return { answer }. Include deterministic data for debugging/future UI.
      return { status: 200, jsonBody: { answer: result.answer, data: result.data } };
    } catch (e) {
      return {
        status: 500,
        jsonBody: { error: "deterministic chat error", details: String(e?.message || e) },
      };
    }
  },
});
