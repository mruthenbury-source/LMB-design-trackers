import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import ChatOverlay from "./ChatOverlay.jsx";

const LS_KEY = "design-programme-workback:v16";

/* ---------- date helpers ---------- */
function isoToday() {
  const d = new Date();
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}
function parseISO(value) {
  if (!value) return null;
  const [y, m, d] = value.split("-").map(Number);
  if (!y || !m || !d) return null;
  const dt = new Date(Date.UTC(y, m - 1, d));
  if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
  return dt;
}
function formatISO(dt) {
  if (!dt) return "";
  const y = dt.getUTCFullYear();
  const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
  const d = String(dt.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}
function addDaysUTC(dt, days) {
  const out = new Date(dt.getTime());
  out.setUTCDate(out.getUTCDate() + days);
  return out;
}
function clampInt(v, fallback = 0) {
  const n = Number(v);
  if (!Number.isFinite(n)) return fallback;
  return Math.trunc(n);
}
function uid() {
  return Math.random().toString(16).slice(2) + Date.now().toString(16);
}
function clean(text) {
  return String(text ?? "").trim();
}
function dayMs() {
  return 24 * 60 * 60 * 1000;
}
function diffDaysUTC(startISO, finishISO) {
  const s = parseISO(startISO);
  const f = parseISO(finishISO);
  if (!s || !f) return null;
  const diff = f.getTime() - s.getTime();
  if (!Number.isFinite(diff)) return null;
  return Math.round(diff / dayMs());
}
/* ---------- UI helpers ---------- */


function datePillStyle({ row, dateKey }) {
  // Not required always grey
  if (row.notRequired) return styles.pillMuted;

  // Completed overrides everything
  if (row.completed) return styles.pillDone;

  // Specific ticked milestones
  if (dateKey === "statusA" && row.statusADone) return styles.pillDone;
  if (dateKey === "firstIssue" && row.firstIssueDone) return styles.pillDone;

  // Late logic (same as tracker)
  if (row.status === "overdue") return styles.pillLate;

  return styles.pillNeutral;
}


/* ---------- schedule model ---------- */
const ANCHORS = [
  { key: "requiredOnSite", label: "Required on Site" },
  { key: "statusA", label: "Status A" },
  { key: "firstIssue", label: "First Issue" },
];

function computeDates({ anchorKey, anchorDateISO, daysReqToStatusA, daysStatusAToFirstIssue }) {
  const anchorDt = parseISO(anchorDateISO);
  if (!anchorDt) return { requiredOnSite: "", statusA: "", firstIssue: "" };

  const d1 = clampInt(daysReqToStatusA, 0);
  const d2 = clampInt(daysStatusAToFirstIssue, 0);

  let req = null;
  let statusA = null;
  let first = null;

  if (anchorKey === "requiredOnSite") {
    req = anchorDt;
    statusA = addDaysUTC(req, -d1);
    first = addDaysUTC(statusA, -d2);
  } else if (anchorKey === "statusA") {
    statusA = anchorDt;
    req = addDaysUTC(statusA, d1);
    first = addDaysUTC(statusA, -d2);
  } else {
    first = anchorDt;
    statusA = addDaysUTC(first, d2);
    req = addDaysUTC(statusA, d1);
  }

  return {
    requiredOnSite: formatISO(req),
    statusA: formatISO(statusA),
    firstIssue: formatISO(first),
  };
}

/* ---------- traffic (Status A) ---------- */
function trafficForRow(row, dates) {
  if (row.notRequired) return "na";
  if (row.completed) return "green";

  const today = parseISO(isoToday());
  const statusA = parseISO(dates.statusA);
  if (!today || !statusA) return "na";

  const daysLeft = Math.ceil((statusA.getTime() - today.getTime()) / dayMs());
  if (daysLeft < 0) return "red";
  if (daysLeft <= 7) return "amber";
  return "green";
}

function TrafficDot({ status, title }) {
  const s =
    status === "red"
      ? { bg: "#EF4444", label: "Overdue (Status A)" }
      : status === "amber"
      ? { bg: "#F59E0B", label: "Due soon (Status A ≤ 7 days)" }
      : status === "green"
      ? { bg: "#10B981", label: "On track (Status A > 7 days)" }
      : { bg: "#9CA3AF", label: "N/A" };

  return (
    <span
      title={title || s.label}
      style={{
        width: 10,
        height: 10,
        borderRadius: 999,
        background: s.bg,
        display: "inline-block",
        flex: "0 0 auto",
      }}
    />
  );
}

function TrafficKeyDotsOnly() {
  return (
    <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
      <span style={styles.keyItem}>
        <TrafficDot status="green" />
        <span style={styles.keyText}>On track</span>
      </span>
      <span style={styles.keyItem}>
        <TrafficDot status="amber" />
        <span style={styles.keyText}>Due soon</span>
      </span>
      <span style={styles.keyItem}>
        <TrafficDot status="red" />
        <span style={styles.keyText}>Overdue</span>
      </span>
      <span style={styles.keyItem}>
        <TrafficDot status="na" />
        <span style={styles.keyText}>N/A</span>
      </span>
    </div>
  );
}

/* ---------- PDF/CSV helpers ---------- */
function escapeHtml(s) {
  return String(s ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
function openPrintWindow({ title, subtitle, tableHead, tableRows }) {
  const w = window.open("", "_blank", "noopener,noreferrer");
  if (!w) return;

  const headHtml = tableHead.map((h) => `<th>${escapeHtml(h)}</th>`).join("");
  const rowsHtml = tableRows
    .map((r) => `<tr>${r.map((c) => `<td>${escapeHtml(c)}</td>`).join("")}</tr>`)
    .join("");

  w.document.write(`
  <!doctype html>
  <html>
    <head>
      <meta charset="utf-8" />
      <title>${escapeHtml(title)}</title>
      <style>
        body { font-family: Arial, sans-serif; padding: 24px; color: #111; }
        h1 { margin: 0 0 6px; font-size: 18px; }
        .sub { margin: 0 0 16px; color: #555; font-size: 12px; }
        table { width: 100%; border-collapse: collapse; font-size: 11px; }
        th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
        th { background: #f5f5f5; text-align: left; }
        @media print { body { padding: 0; } }
      </style>
    </head>
    <body>
      <h1>${escapeHtml(title)}</h1>
      <div class="sub">${escapeHtml(subtitle)}</div>
      <table>
        <thead><tr>${headHtml}</tr></thead>
        <tbody>${rowsHtml}</tbody>
      </table>
      <script>window.onload = () => { window.print(); };</script>
    </body>
  </html>
  `);
  w.document.close();
}
function toCSVCell(x) {
  const s = String(x ?? "");
  const needsQuotes = /[",\n]/.test(s);
  return needsQuotes ? `"${s.replaceAll('"', '""')}"` : s;
}
function downloadCSV(filename, lines) {
  const blob = new Blob([lines.join("\n")], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

/* ---------- defaults ---------- */
function defaultLevel(name = "Level 1") {
  return { id: uid(), name, startDate: "", finishDate: "" };
}
function defaultMasterRow() {
  return { id: uid(), blockZone: "", levels: [defaultLevel("Level 1")] };
}
function defaultResponsibility() {
  return { id: uid(), name: "", supplier: "" };
}
function defaultRow(kind = "item") {
  return {
    id: uid(),
    kind, // "header" | "item"
    item: "",
    anchorKey: "requiredOnSite",
    anchorDateISO: "",
    overrideDaysReqToStatusA: null,
    overrideDaysStatusAToFirstIssue: null,
    completed: false,
    notRequired: false,
    statusADone: false,
    firstIssueDone: false,
    // locks are set when a guest ticks a milestone (admin can unlock)
    locks: { statusADone: null, firstIssueDone: null },
    // comments are append-only (each saved comment is locked)
    comments: [],
    meta: { generated: false, blockZone: "", levelId: null, levelName: "", finishDate: "" },
  };
}
function defaultPage(name = "Page") {
  return { id: uid(), name, rows: [], meta: { generated: false, responsibilityId: null, isMaster: false } };
}
function defaultProject(name = "New Project") {
  return {
    id: uid(),
    name,
    master: [defaultMasterRow()],
    responsibilities: [defaultResponsibility()],
    pages: [
      {
        id: uid(),
        name: "Project Home",
        rows: [],
        meta: { generated: false, responsibilityId: null, isMaster: true },
      },
    ],
  };
}

/* ---------- build generated rows from master schedule ---------- */
function buildGeneratedRows(pageName, master) {
  const rows = [];
  const page = clean(pageName);

  (master || []).forEach((m) => {
    const bz = clean(m.blockZone);
    const levels = Array.isArray(m.levels) ? m.levels : [];
    if (!bz || levels.length === 0) return;

    const header = defaultRow("header");
    header.item = bz;
    header.meta = { generated: true, blockZone: bz, levelId: null, levelName: "", finishDate: "" };
    rows.push(header);

    levels.forEach((lv, idx) => {
      const levelName = clean(lv?.name) || `Level ${idx + 1}`;
      const r = defaultRow("item");
      r.item = `${page}_${bz}_${levelName}`;
      r.meta = { generated: true, blockZone: bz, levelId: lv?.id || null, levelName, finishDate: lv?.finishDate || "" };

      r.anchorKey = "requiredOnSite";
      r.anchorDateISO = lv?.startDate || "";

      rows.push(r);
    });
  });

  return rows;
}

function overdueFlags(row, dates, todayISO) {
  if (row.notRequired) return { overdue: false, overdueReq: false, overdueA: false, overdueF: false };
  const today = parseISO(todayISO);
  if (!today) return { overdue: false, overdueReq: false, overdueA: false, overdueF: false };

  const req = parseISO(dates.requiredOnSite);
  const a = parseISO(dates.statusA);
  const f = parseISO(dates.firstIssue);

  const overdueReq = !!req && req.getTime() < today.getTime() && !row.completed;
  const overdueA = !!a && a.getTime() < today.getTime() && !row.statusADone && !row.completed;
  const overdueF = !!f && f.getTime() < today.getTime() && !row.firstIssueDone && !row.completed;

  return { overdue: overdueReq || overdueA || overdueF, overdueReq, overdueA, overdueF };
}

/* ---------- Programme (Gantt) ---------- */
function buildProgrammeItems(master) {
  const items = [];
  (master || []).forEach((m) => {
    const bz = clean(m.blockZone);
    (m.levels || []).forEach((lv, idx) => {
      const levelName = clean(lv?.name) || `Level ${idx + 1}`;
      const startISO = lv?.startDate || "";
      const finishISO = lv?.finishDate || lv?.startDate || "";
      const start = parseISO(startISO);
      const finish = parseISO(finishISO);
      items.push({
        id: lv?.id || uid(),
        label: bz ? `${bz} — ${levelName}` : levelName,
        startISO,
        finishISO,
        start,
        finish,
      });
    });
  });
  return items;
}

function ProgrammeGantt({ master, dense = false }) {
  const items = useMemo(() => buildProgrammeItems(master), [master]);

  const range = useMemo(() => {
    const valid = items
      .map((i) => ({ s: i.start, f: i.finish }))
      .filter((x) => x.s && x.f)
      .map((x) => ({ s: x.s.getTime(), f: x.f.getTime() }));

    if (!valid.length) return null;

    let min = valid[0].s;
    let max = valid[0].f;
    for (const v of valid) {
      if (v.s < min) min = v.s;
      if (v.f > max) max = v.f;
    }
    min -= 3 * dayMs();
    max += 3 * dayMs();

    const total = Math.max(1, max - min);
    return { min, max, total };
  }, [items]);

  const monthTicks = useMemo(() => {
    if (!range) return [];
    const start = new Date(range.min);
    const end = new Date(range.max);

    const cur = new Date(Date.UTC(start.getUTCFullYear(), start.getUTCMonth(), 1));
    const ticks = [];
    while (cur.getTime() <= end.getTime()) {
      const t = cur.getTime();
      const pct = ((t - range.min) / range.total) * 100;
      ticks.push({
        key: `${cur.getUTCFullYear()}-${cur.getUTCMonth() + 1}`,
        pct,
        label: cur.toLocaleString(undefined, { month: "short", year: "2-digit" }),
      });
      cur.setUTCMonth(cur.getUTCMonth() + 1);
    }
    return ticks;
  }, [range]);

  if (!items.length) return <div style={styles.muted}>Add Blocks/Zones and Levels to see the programme.</div>;
  if (!range) return <div style={styles.muted}>Add start/finish dates to see the programme.</div>;

  const rowStyle = dense ? styles.ganttRowDense : styles.ganttRow;
  const headerStyle = dense ? styles.ganttHeaderRowDense : styles.ganttHeaderRow;

  return (
    <div style={styles.ganttWrap}>
      <div style={headerStyle}>
        <div style={styles.ganttLabelCol}>Block/Zone — Level</div>
        <div style={styles.ganttChartCol}>
          <div style={styles.ganttAxis}>
            {monthTicks.map((t) => (
              <div key={t.key} style={{ ...styles.ganttTick, left: `${t.pct}%` }}>
                <div style={styles.ganttTickLine} />
                <div style={styles.ganttTickLabel}>{t.label}</div>
              </div>
            ))}
          </div>
        </div>
        <div style={styles.ganttDateCol}>Start</div>
        <div style={styles.ganttDateCol}>Finish</div>
      </div>

      {items.map((it) => {
        const s = it.start ? it.start.getTime() : null;
        const f = it.finish ? it.finish.getTime() : null;

        const hasDates = s != null && f != null;
        const startPct = hasDates ? ((s - range.min) / range.total) * 100 : 0;
        const endPct = hasDates ? ((f - range.min) / range.total) * 100 : 0;

        const left = Math.max(0, Math.min(100, startPct));
        const right = Math.max(0, Math.min(100, endPct));
        const width = Math.max(1, right - left);

        return (
          <div key={it.id} style={rowStyle}>
            <div style={styles.ganttLabelCol}>
              <div style={{ fontSize: 12, color: "#111827", lineHeight: 1.25 }}>{it.label || "—"}</div>
            </div>

            <div style={styles.ganttChartCol}>
              <div style={styles.ganttLane}>
                {hasDates ? (
                  <div style={{ ...styles.ganttBar, left: `${left}%`, width: `${width}%` }} title={`${it.startISO || "—"} → ${it.finishISO || "—"}`} />
                ) : (
                  <div style={styles.ganttBarMissing}>No dates</div>
                )}
              </div>
            </div>

            <div style={styles.ganttDateCol}>
              <div style={styles.pillCompact}>{it.startISO || "—"}</div>
            </div>
            <div style={styles.ganttDateCol}>
              <div style={styles.pillCompact}>{it.finishISO || "—"}</div>
            </div>
          </div>
        );
      })}
    </div>
  );
}

/* ---------- app pages ---------- */
const VIEW = {
  LANDING: "landing",
  PROJECT: "project",
  SUMMARY: "summary",
  GANTT_SUMMARY: "gantt_summary",
};

export default function App() {
  // Boot / hydration
  const [isHydrating, setIsHydrating] = useState(true);
  const [hydrateError, setHydrateError] = useState(null);

  // ✅ Blob save status (auto-save for all users)
  const [saveStatus, setSaveStatus] = useState("idle"); // idle | saving | saved | error
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [saveError, setSaveError] = useState(null);
  const saveTimerRef = useRef(null);

  const [globalDaysReqToStatusA, setGlobalDaysReqToStatusA] = useState(14);
  const [globalDaysStatusAToFirstIssue, setGlobalDaysStatusAToFirstIssue] = useState(28);

  const [projects, setProjects] = useState([defaultProject("Project 1")]);
  const [activeProjectId, setActiveProjectId] = useState(null);
  const [activePageId, setActivePageId] = useState(null);

  const [view, setView] = useState(VIEW.LANDING);
  const didHydrateRef = useRef(false);

  /* ---------------- AUTH + PERMISSIONS (Azure Static Web Apps) ---------------- */
  const [authUser, setAuthUser] = useState(undefined); // undefined=loading, null=anonymous, object=logged in

  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const r = await fetch("/.auth/me", { cache: "no-store" });
        if (!r.ok) {
          if (!cancelled) setAuthUser(null);
          return;
        }
        const data = await r.json();
        const u = data?.clientPrincipal || null;
        if (!cancelled) setAuthUser(u);
      } catch {
        if (!cancelled) setAuthUser(null);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  const authRoles = useMemo(() => {
    if (!authUser) return [];
    return Array.isArray(authUser?.userRoles) ? authUser.userRoles : [];
  }, [authUser]);

  const isAdmin = useMemo(() => {
    // Treat anonymous as admin so the site still works without login
    if (authUser === null) return true;
    if (!authUser) return false;
    const roles = authRoles.map((r) => String(r || "").toLowerCase());
    return roles.includes("administrator") || roles.includes("admin");
  }, [authUser, authRoles]);

  const isGuest = !!authUser && !isAdmin;

  const hasSupplierAccess = useCallback(
    (supplier) => {
      if (isAdmin) return true;
      const s = String(supplier || "").trim().toLowerCase();
      if (!s) return false;
      const roles = authRoles.map((r) => String(r || "").trim().toLowerCase());
      return roles.includes(s) || roles.includes(`supplier:${s}`);
    },
    [isAdmin, authRoles]
  );

  // Guests only see projects that have at least one tracker page for their supplier
  const visibleProjects = useMemo(() => {
    if (isAdmin) return projects;
    if (!projects?.length) return [];
    return projects.filter((p) =>
      (p.pages || []).some((pg) => {
        if (pg.meta?.isMaster) return false;
        const respId = pg.meta?.responsibilityId;
        const resp = (p.responsibilities || []).find((r) => r.id === respId);
        const supplier = String(resp?.supplier || "").trim();
        return supplier && hasSupplierAccess(supplier);
      })
    );
  }, [projects, isAdmin, hasSupplierAccess]);

  // Summary filters
  const [summaryFilter, setSummaryFilter] = useState("ongoing");
  const [summaryProjectId, setSummaryProjectId] = useState("all");
  const [summarySupplier, setSummarySupplier] = useState("all");

  // Programme Summary print target
  const programmePrintRef = useRef(null);

  /* ---------------- CHATBOT (UI + DATA CONTEXT) ---------------- */
  const [chatOpen, setChatOpen] = useState(false);
  const [chatMessages, setChatMessages] = useState([
    {
      role: "assistant",
      content:
        "Ask me about your programme data. Examples:\n• What is overdue next?\n• Which supplier has the most overdue items?\n• What are the next Status A dates in Project 1?",
    },
  ]);
  const [chatInput, setChatInput] = useState("");
  const [chatBusy, setChatBusy] = useState(false);
  const [chatStatus, setChatStatus] = useState("idle"); // "idle" | "checking" | "ok" | "error"
  const [chatSearchBackups, setChatSearchBackups] = useState(false);
  const chatEndRef = useRef(null);
  const CHAT_WELCOME = {
    role: "assistant",
    content:
      "Ask me about your programme data. Examples:\n• What is overdue next?\n• Which supplier has the most overdue items?\n• What are the next Status A dates in Project 1?",
  };

  function resetChat() {
    setChatMessages([CHAT_WELCOME]);
    setChatInput("");
    setChatBusy(false);
    setChatStatus("idle");
  }


  useEffect(() => {
    if (!chatOpen) return;
    const t = setTimeout(() => {
      try {
        chatEndRef.current?.scrollIntoView({ behavior: "smooth", block: "end" });
      } catch {
        // ignore
      }
    }, 0);
    return () => clearTimeout(t);
  }, [chatOpen, chatMessages]);

  // Shared hydration routine (used for Blob + localStorage fallback)
  async function hydrateFromPayload(parsed) {
    if (!parsed) return;

      if (Number.isFinite(parsed.globalDaysReqToStatusA)) setGlobalDaysReqToStatusA(parsed.globalDaysReqToStatusA);
      if (Number.isFinite(parsed.globalDaysStatusAToFirstIssue)) setGlobalDaysStatusAToFirstIssue(parsed.globalDaysStatusAToFirstIssue);
  
      if (Array.isArray(parsed.projects) && parsed.projects.length) {
        // your hydration logic (keep as-is)
        const hydrated = parsed.projects.map((proj) => {
          const master =
            Array.isArray(proj.master) && proj.master.length
              ? proj.master.map((m) => ({
                  id: m.id || uid(),
                  blockZone: m.blockZone || "",
                  levels:
                    Array.isArray(m.levels) && m.levels.length
                      ? m.levels.map((lv, idx) => ({
                          id: lv.id || uid(),
                          name: lv.name || `Level ${idx + 1}`,
                          startDate: lv.startDate || "",
                          finishDate: lv.finishDate || "",
                        }))
                      : [defaultLevel("Level 1")],
                }))
              : [defaultMasterRow()];
  
          const responsibilities =
            Array.isArray(proj.responsibilities) && proj.responsibilities.length
              ? proj.responsibilities.map((r) => ({
                  id: r.id || uid(),
                  name: r.name || "",
                  supplier: r.supplier || "",
                }))
              : [defaultResponsibility()];
  
          const pages =
            Array.isArray(proj.pages) && proj.pages.length
              ? proj.pages.map((pg) => {
                  const rawName = pg.name || "Untitled Page";
                  const migratedName = pg.meta?.isMaster && rawName === "Master" ? "Project Home" : rawName;
  
                  return {
                    id: pg.id || uid(),
                    name: migratedName,
                    rows: Array.isArray(pg.rows)
                      ? pg.rows.map((r) => ({
                          ...defaultRow(r.kind || "item"),
                          ...r,
                          id: r.id || uid(),
                          kind: r.kind || "item",
                          completed: !!r.completed,
                          notRequired: !!r.notRequired,
                          statusADone: !!r.statusADone,
                          firstIssueDone: !!r.firstIssueDone,
                          locks: {
                            statusADone: r?.locks?.statusADone || null,
                            firstIssueDone: r?.locks?.firstIssueDone || null,
                          },
                          comments: Array.isArray(r?.comments) ? r.comments : [],
                          meta: {
                            generated: !!r?.meta?.generated,
                            blockZone: r?.meta?.blockZone || "",
                            levelId: r?.meta?.levelId ?? null,
                            levelName: r?.meta?.levelName || "",
                            finishDate: r?.meta?.finishDate || "",
                          },
                        }))
                      : [],
                    meta: {
                      generated: !!pg?.meta?.generated,
                      responsibilityId: pg?.meta?.responsibilityId ?? null,
                      isMaster: !!pg?.meta?.isMaster,
                    },
                  };
                })
              : [];
  
          if (!pages.some((p) => p.meta?.isMaster)) {
            pages.unshift({
              id: uid(),
              name: "Project Home",
              rows: [],
              meta: { generated: false, responsibilityId: null, isMaster: true },
            });
          }
  
          return { id: proj.id || uid(), name: proj.name || "Untitled Project", master, responsibilities, pages };
        });
  
        setProjects(hydrated);
  
        const pid = parsed.activeProjectId || hydrated[0].id;
        setActiveProjectId(pid);
  
        const proj0 = hydrated.find((p) => p.id === pid) || hydrated[0];
        const masterPg = proj0.pages.find((p) => p.meta?.isMaster) || proj0.pages[0];
        setActivePageId(parsed.activePageId || masterPg.id);
      }
  
      if (parsed.view && Object.values(VIEW).includes(parsed.view)) setView(parsed.view);
      if (typeof parsed.summaryFilter === "string") setSummaryFilter(parsed.summaryFilter);
      if (typeof parsed.summaryProjectId === "string") setSummaryProjectId(parsed.summaryProjectId);
      if (typeof parsed.summarySupplier === "string") setSummarySupplier(parsed.summarySupplier);
  }

  /* ---- load (Blob first) ---- */
  useEffect(() => {
    let cancelled = false;
    (async () => {
      setIsHydrating(true);
      setHydrateError(null);

      try {
        // 1) Prefer Blob-backed state via API
        const res = await fetch("/api/state", { method: "GET", cache: "no-store" });
        if (!res.ok) throw new Error(`State load failed (${res.status})`);
        const data = await res.json().catch(() => ({}));
        const serverState = data?.state ?? null;

        if (!cancelled && serverState) {
          await hydrateFromPayload(serverState);
          try {
            localStorage.setItem(LS_KEY, JSON.stringify(serverState));
          } catch {}
          return;
        }
      } catch (e) {
        if (!cancelled) setHydrateError(String(e?.message || e));
      }

      // 2) Fallback to localStorage
      try {
        const raw = localStorage.getItem(LS_KEY);
        const parsed = raw ? JSON.parse(raw) : null;
        if (!cancelled && parsed) await hydrateFromPayload(parsed);
      } catch {
        // ignore
      }
    })()
      .catch(() => {
        // ignore
      })
      .finally(() => {
        if (cancelled) return;
        didHydrateRef.current = true;
        setIsHydrating(false);
      });

    return () => {
      cancelled = true;
    };
  }, []);
  


  // ✅ Save state to Azure Blob via API
  const saveToBlobNow = useCallback(async (statePayload) => {
    setSaveStatus("saving");
    setSaveError(null);
    try {
      const res = await fetch("/api/state", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ state: statePayload }),
      });
      if (!res.ok) {
        const txt = await res.text().catch(() => "");
        throw new Error(`POST /api/state failed (${res.status}) ${txt || ""}`.trim());
      }
      setSaveStatus("saved");
      setLastSavedAt(new Date().toISOString());
    } catch (e) {
      setSaveStatus("error");
      setSaveError(String(e?.message || e));
    }
  }, []);

  /* ---- persist ---- */
  useEffect(() => {
    // ✅ prevent overwriting saved data on the first render
    if (!didHydrateRef.current) return;
  
    const payload = {
      globalDaysReqToStatusA,
      globalDaysStatusAToFirstIssue,
      projects,
      activeProjectId,
      activePageId,
      view,
      summaryFilter,
      summaryProjectId,
      summarySupplier,
    };
  
    try {
      localStorage.setItem(LS_KEY, JSON.stringify(payload));

    // ✅ Debounced auto-save to Blob for all users
    clearTimeout(saveTimerRef.current);
    saveTimerRef.current = setTimeout(() => {
      saveToBlobNow(payload);
    }, 700);
    } catch {
      // ignore
    }
  }, [
    globalDaysReqToStatusA,
    globalDaysStatusAToFirstIssue,
    projects,
    activeProjectId,
    activePageId,
    view,
    summaryFilter,
    summaryProjectId,
    summarySupplier,
  ]);
  
  /* ---- derived active project/page ---- */
  const projectPool = useMemo(() => {
    if (isAdmin) return projects;
    // while auth is loading, keep current behaviour
    if (authUser === undefined) return projects;
    return visibleProjects?.length ? visibleProjects : projects;
  }, [projects, visibleProjects, isAdmin, authUser]);

  const activeProject = useMemo(
    () => projectPool.find((p) => p.id === (activeProjectId || projectPool[0]?.id)) || projectPool[0] || null,
    [projectPool, activeProjectId]
  );

  const activePage = useMemo(() => {
    if (!activeProject) return null;
    return (
      activeProject.pages.find((pg) => pg.id === (activePageId || activeProject.pages[0]?.id)) ||
      activeProject.pages[0] ||
      null
    );
  }, [activeProject, activePageId]);

  // Ensure we always have valid active IDs
  useEffect(() => {
    if (!projectPool.length) return;
    if (!activeProjectId) setActiveProjectId(projectPool[0].id);
  }, [projectPool, activeProjectId]);

  useEffect(() => {
    if (!activeProject) return;
    if (!activePageId) {
      const mp = activeProject.pages.find((p) => p.meta?.isMaster) || activeProject.pages[0];
      setActivePageId(mp?.id || null);
    }
  }, [activeProject, activePageId]);

  // Guest guardrails: always land on an allowed tracker page, and prevent navigating elsewhere
  useEffect(() => {
    if (!isGuest) return;
    if (authUser === undefined) return; // wait for auth probe
    if (!visibleProjects?.length) return;

    if (view !== VIEW.PROJECT) setView(VIEW.PROJECT);

    const projOk = visibleProjects.some((p) => p.id === activeProjectId);
    const proj = (projOk ? visibleProjects.find((p) => p.id === activeProjectId) : visibleProjects[0]) || null;
    if (!proj) return;
    if (!projOk) setActiveProjectId(proj.id);

    const allowed = (proj.pages || []).filter((pg) => {
      if (pg.meta?.isMaster) return false;
      const respId = pg.meta?.responsibilityId;
      const resp = (proj.responsibilities || []).find((r) => r.id === respId);
      const supplier = String(resp?.supplier || "").trim();
      return supplier && hasSupplierAccess(supplier);
    });

    if (!allowed.length) {
      if (activePageId !== null) setActivePageId(null);
      return;
    }

    const pageOk = allowed.some((pg) => pg.id === activePageId);
    if (!pageOk) setActivePageId(allowed[0].id);
  }, [isGuest, authUser, visibleProjects, activeProjectId, activePageId, view, hasSupplierAccess]);

  /* ---- update helpers ---- */
  function updateProject(projectId, patch) {
    setProjects((prev) => prev.map((p) => (p.id === projectId ? { ...p, ...patch } : p)));
  }

  function updatePage(projectId, pageId, patch) {
    setProjects((prev) =>
      prev.map((p) =>
        p.id !== projectId
          ? p
          : {
              ...p,
              pages: p.pages.map((pg) => (pg.id === pageId ? { ...pg, ...patch } : pg)),
            }
      )
    );
  }

  function updateRow(rowId, patch) {
    if (!activeProject || !activePage) return;

    setProjects((prev) =>
      prev.map((p) => {
        if (p.id !== activeProject.id) return p;

        return {
          ...p,
          pages: p.pages.map((pg) =>
            pg.id !== activePage.id
              ? pg
              : {
                  ...pg,
                  rows: (pg.rows || []).map((r) => (r.id === rowId ? { ...r, ...patch } : r)),
                }
          ),
        };
      })
    );
  }

  /* ---- master edits ---- */
  function addMasterRow() {
    if (!activeProject) return;
    updateProject(activeProject.id, { master: [...(activeProject.master || []), defaultMasterRow()] });
  }
  function updateMasterRow(masterId, patch) {
    if (!activeProject) return;
    updateProject(activeProject.id, {
      master: (activeProject.master || []).map((m) => (m.id === masterId ? { ...m, ...patch } : m)),
    });
  }
  function removeMasterRow(masterId) {
    if (!activeProject) return;
    const next = (activeProject.master || []).filter((m) => m.id !== masterId);
    updateProject(activeProject.id, { master: next.length ? next : [defaultMasterRow()] });
  }
  function addLevel(masterId) {
    if (!activeProject) return;
    const next = (activeProject.master || []).map((m) => {
      if (m.id !== masterId) return m;
      const current = Array.isArray(m.levels) ? m.levels : [];
      const name = `Level ${current.length + 1}`;
      return { ...m, levels: [...current, defaultLevel(name)] };
    });
    updateProject(activeProject.id, { master: next });
  }
  function updateLevel(masterId, levelId, patch) {
    if (!activeProject) return;
    const next = (activeProject.master || []).map((m) => {
      if (m.id !== masterId) return m;
      const levels = (m.levels || []).map((lv) => (lv.id === levelId ? { ...lv, ...patch } : lv));
      return { ...m, levels };
    });
    updateProject(activeProject.id, { master: next });
  }
  function removeLevel(masterId, levelId) {
    if (!activeProject) return;
    const next = (activeProject.master || []).map((m) => {
      if (m.id !== masterId) return m;
      const levels = (m.levels || []).filter((lv) => lv.id !== levelId);
      return { ...m, levels: levels.length ? levels : [defaultLevel("Level 1")] };
    });
    updateProject(activeProject.id, { master: next });
  }

  function addResponsibility() {
    if (!activeProject) return;
    updateProject(activeProject.id, {
      responsibilities: [...(activeProject.responsibilities || []), defaultResponsibility()],
    });
  }
  function updateResponsibility(respId, patch) {
    if (!activeProject) return;
    updateProject(activeProject.id, {
      responsibilities: (activeProject.responsibilities || []).map((r) => (r.id === respId ? { ...r, ...patch } : r)),
    });
  }
  function removeResponsibility(respId) {
    if (!activeProject) return;
    const remaining = (activeProject.responsibilities || []).filter((r) => r.id !== respId);
    updateProject(activeProject.id, { responsibilities: remaining.length ? remaining : [defaultResponsibility()] });
  }

  /* ---- auto-create pages from responsibilities ---- */
  useEffect(() => {
    if (!activeProject) return;

    const responsibilities = (activeProject.responsibilities || [])
      .map((r) => ({ ...r, name: clean(r.name) }))
      .filter((r) => r.name.length > 0);

    const masterPage = activeProject.pages.find((p) => p.meta?.isMaster);
    const existingPages = activeProject.pages || [];

    const nextPages = [];
    if (masterPage) nextPages.push(masterPage);

    const existingGeneratedByResp = new Map(
      existingPages
        .filter((p) => p.meta?.generated && p.meta?.responsibilityId)
        .map((p) => [p.meta.responsibilityId, p])
    );

    responsibilities.forEach((r) => {
      const existing = existingGeneratedByResp.get(r.id);
      if (existing) nextPages.push(existing.name !== r.name ? { ...existing, name: r.name } : existing);
      else nextPages.push({ ...defaultPage(r.name), meta: { generated: true, responsibilityId: r.id, isMaster: false } });
    });

    existingPages.filter((p) => !p.meta?.isMaster && !p.meta?.generated).forEach((p) => nextPages.push(p));

    const sig = (pages) =>
      pages
        .map((p) => `${p.meta?.isMaster ? "M" : p.meta?.generated ? "G" : "U"}:${p.id}:${p.name}:${p.meta?.responsibilityId ?? ""}`)
        .join("|");
    if (sig(existingPages) === sig(nextPages)) return;

    updateProject(activeProject.id, { pages: nextPages });

    if (!nextPages.some((p) => p.id === activePageId)) {
      const mp = nextPages.find((p) => p.meta?.isMaster) || nextPages[0];
      setActivePageId(mp?.id || null);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeProject?.id, activeProject?.responsibilities]);

  /* ---- auto-generate tracker rows from master for non-master pages ---- */
  useEffect(() => {
    if (!activeProject || !activePage) return;
    if (activePage.meta?.isMaster) return;

    const generated = buildGeneratedRows(activePage.name, activeProject.master);
    const manual = (activePage.rows || []).filter((r) => !r?.meta?.generated);
    const nextRows = [...generated, ...manual];

    const sig = (rows) => rows.map((r) => `${r.kind}:${r.item}:${r.meta?.blockZone || ""}:${r.meta?.levelId || ""}`).join("|");
    if (sig(activePage.rows || []) === sig(nextRows)) return;

    updatePage(activeProject.id, activePage.id, { rows: nextRows });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [activeProject?.id, activeProject?.master, activePage?.id, activePage?.name, activePage?.meta?.isMaster]);

  function addManualRow() {
    if (!activeProject || !activePage || activePage.meta?.isMaster) return;
    const r = defaultRow("item");
    r.meta.generated = false;
    updatePage(activeProject.id, activePage.id, { rows: [...activePage.rows, r] });
  }
  function removeRow(rowId) {
    if (!activeProject || !activePage || activePage.meta?.isMaster) return;
    updatePage(activeProject.id, activePage.id, { rows: activePage.rows.filter((r) => r.id !== rowId) });
  }

  function addProject() {
    const p = defaultProject(`Project ${projects.length + 1}`);
    setProjects((prev) => [...prev, p]);
    setActiveProjectId(p.id);
    const mp = p.pages.find((x) => x.meta?.isMaster) || p.pages[0];
    setActivePageId(mp.id);
    setView(VIEW.PROJECT);
  }

  function goToProject() {
    if (!activeProject) return;
    const mp = activeProject.pages.find((x) => x.meta?.isMaster) || activeProject.pages[0];
    setActivePageId(mp.id);
    setView(VIEW.PROJECT);
  }

  /* ---- computed rows ---- */
  const computedRows = useMemo(() => {
    const rows = activePage?.rows || [];
    const todayISO = isoToday();

    return rows.map((r) => {
      const d1 = r.overrideDaysReqToStatusA ?? globalDaysReqToStatusA;
      const d2 = r.overrideDaysStatusAToFirstIssue ?? globalDaysStatusAToFirstIssue;

      const dates =
        r.kind === "header"
          ? { requiredOnSite: "", statusA: "", firstIssue: "" }
          : computeDates({
              anchorKey: r.anchorKey,
              anchorDateISO: r.anchorDateISO,
              daysReqToStatusA: d1,
              daysStatusAToFirstIssue: d2,
            });

      const o = r.kind === "header" ? { overdue: false, overdueReq: false, overdueA: false, overdueF: false } : overdueFlags(r, dates, todayISO);
      const traffic = r.kind === "header" ? "na" : trafficForRow(r, dates);

      return { ...r, _computed: dates, _d1: d1, _d2: d2, _overdue: o, _traffic: traffic };
    });
  }, [activePage, globalDaysReqToStatusA, globalDaysStatusAToFirstIssue]);

  /* ---- summary aggregation ---- */
  const summaryItems = useMemo(() => {
    const todayISO = isoToday();
    const out = [];

    projects.forEach((proj) => {
      const supplierByRespId = new Map((proj.responsibilities || []).map((r) => [r.id, r.supplier || ""]));

      proj.pages
        .filter((pg) => !pg?.meta?.isMaster)
        .forEach((pg) => {
          (pg.rows || [])
            .filter((r) => r.kind !== "header")
            .filter((r) => !r.notRequired)
            .forEach((r) => {
              const d1 = r.overrideDaysReqToStatusA ?? globalDaysReqToStatusA;
              const d2 = r.overrideDaysStatusAToFirstIssue ?? globalDaysStatusAToFirstIssue;

              const dates = computeDates({
                anchorKey: r.anchorKey,
                anchorDateISO: r.anchorDateISO,
                daysReqToStatusA: d1,
                daysStatusAToFirstIssue: d2,
              });

              const o = overdueFlags(r, dates, todayISO);
              const status = r.completed ? "done" : o.overdue ? "overdue" : "ongoing";
              const traffic = trafficForRow(r, dates);

              const comments = Array.isArray(r.comments) ? r.comments : [];
              const commentsText = comments
                .map((c) => {
                  const who = clean(c?.name) || clean(c?.lockedBy) || "";
                  const when = clean(c?.dateISO) || "";
                  const txt = clean(c?.text) || "";
                  return [when, who, txt].filter(Boolean).join(" — ");
                })
                .filter(Boolean)
                .join("\n");

              const totalDurationDays = diffDaysUTC(dates.firstIssue, dates.requiredOnSite);

              out.push({
              projectId: proj.id,
              projectName: proj.name,
              pageId: pg.id,
              pageName: pg.name,
              rowId: r.id,
              title: r.item || "Unknown",
              supplier: supplierByRespId.get(pg.meta?.responsibilityId) || "Unknown",
              requiredOnSite: dates.requiredOnSite || "",
              statusA: dates.statusA || "",
              firstIssue: dates.firstIssue || "",
              anchorKey: r.anchorKey || "N/A",
              anchorDateISO: r.anchorDateISO || "N/A",
              daysReqToStatusA: d1 || 0,
              daysStatusAToFirstIssue: d2 || 0,
              totalDurationDays: totalDurationDays || 0,
              timeframe: dates.firstIssue && dates.requiredOnSite ? `${dates.firstIssue} → ${dates.requiredOnSite}` : "",
              commentsText: commentsText || "",
              commentsCount: comments.length || 0,

              // tickboxes
              completed: !!r.completed,
              notRequired: !!r.notRequired,
              statusADone: !!r.statusADone,
              firstIssueDone: !!r.firstIssueDone,

              // ✅ canonical approval state
              approvalState: {
              statusA: !!r.statusADone,
              firstIssue: !!r.firstIssueDone,
              completed: !!r.completed,
              overdue: status === "overdue",
              },
                searchText: [
                proj.name,
                pg.name,
                r.item,
                supplierByRespId.get(pg.meta?.responsibilityId),
                `statusA:${!!r.statusADone}`,
                `firstIssue:${!!r.firstIssueDone}`,
                `completed:${!!r.completed}`,
                `overdue:${status === "overdue"}`,
                ].filter(Boolean).join(" | "),

                status,
              traffic,
              });
            });
        });
    });


    const rank = { overdue: 0, ongoing: 1, done: 2 };
    out.sort((a, b) => {
      const ra = rank[a.status] ?? 9;
      const rb = rank[b.status] ?? 9;
      if (ra !== rb) return ra - rb;
      const da = parseISO(a.statusA)?.getTime() ?? Number.MAX_SAFE_INTEGER;
      const db = parseISO(b.statusA)?.getTime() ?? Number.MAX_SAFE_INTEGER;
      if (da !== db) return da - db;
      return a.title.localeCompare(b.title);
    });

    return out;
  }, [projects, globalDaysReqToStatusA, globalDaysStatusAToFirstIssue]);

  const supplierOptions = useMemo(() => {
    const set = new Set();
    projects.forEach((p) =>
      (p.responsibilities || []).forEach((r) => {
        const s = clean(r.supplier);
        if (s) set.add(s);
      })
    );
    return Array.from(set).sort((a, b) => a.localeCompare(b));
  }, [projects]);

  const filteredSummary = useMemo(() => {
    let arr = summaryItems;
    if (summaryFilter !== "all") arr = arr.filter((x) => x.status === summaryFilter);
    if (summaryProjectId !== "all") arr = arr.filter((x) => x.projectId === summaryProjectId);
    if (summarySupplier !== "all") arr = arr.filter((x) => x.supplier === summarySupplier);
    return arr;
  }, [summaryItems, summaryFilter, summaryProjectId, summarySupplier]);

  function jumpToItem(item) {
    setActiveProjectId(item.projectId);
    setActivePageId(item.pageId);
    setView(VIEW.PROJECT);
  }

  function exportSummaryCSV() {
    const header = ["Status", "Traffic", "Project", "Responsibility", "Supplier", "Item", "Required on Site", "Status A", "First Issue"];
    const lines = [header.join(",")];

    filteredSummary.forEach((r) => {
      lines.push(
        [
          toCSVCell(r.status),
          toCSVCell(r.traffic),
          toCSVCell(r.projectName),
          toCSVCell(r.pageName),
          toCSVCell(r.supplier),
          toCSVCell(r.title),
          toCSVCell(r.requiredOnSite),
          toCSVCell(r.statusA),
          toCSVCell(r.firstIssue),
        ].join(",")
      );
    });

    downloadCSV(`Summary-${summaryFilter}-${isoToday()}.csv`.replace(/\s+/g, "-"), lines);
  }

  function exportSummaryPDF() {
    openPrintWindow({
      title: `LMB Design Programme and Trackers — Summary (${summaryFilter.toUpperCase()})`,
      subtitle: `Generated: ${isoToday()} • Traffic based on Status A date`,
      tableHead: ["Status", "Traffic", "Project", "Responsibility", "Supplier", "Item", "Req. on Site", "Status A", "First Issue"],
      tableRows: filteredSummary.map((r) => [
        r.status,
        r.traffic,
        r.projectName,
        r.pageName,
        r.supplier || "",
        r.title,
        r.requiredOnSite || "",
        r.statusA || "",
        r.firstIssue || "",
      ]),
    });
  }

  const isMasterPage = !!activePage?.meta?.isMaster;

  /* ---------- PRINT PROGRAMME SUMMARY (only) ---------- */
  function printProgrammeSummary() {
    const node = programmePrintRef.current;
    if (!node) return;

    const html = node.innerHTML;

    const w = window.open("", "_blank", "noopener,noreferrer");
    if (!w) return;

    w.document.write(`
      <!doctype html>
      <html>
        <head>
          <meta charset="utf-8"/>
          <title>Programme Summary</title>
          <style>
            @page { margin: 12mm; }
            body { font-family: Arial, sans-serif; color: #111; margin: 0; padding: 0; }
            .wrap { padding: 12mm; }
            h1 { font-size: 18px; margin: 0 0 6px; }
            .sub { font-size: 12px; color: #555; margin: 0 0 14px; }
            .project { margin: 0 0 18px; page-break-inside: avoid; }
            .project h2 { font-size: 14px; margin: 0 0 6px; }
            .ganttWrap { display: grid; gap: 8px; }
            .ganttHeaderRow, .ganttRow { display: grid; grid-template-columns: 260px 1fr 120px 120px; gap: 10px; align-items: center; }
            .ganttHeaderRow { align-items: end; }
            .ganttAxis { position: relative; height: 34px; border: 1px solid #e5e7eb; border-radius: 10px; background: #f9fafb; overflow: hidden; }
            .ganttTick { position: absolute; top: 0; bottom: 0; transform: translateX(-0.5px); }
            .ganttTickLine { position: absolute; top: 0; bottom: 0; width: 1px; background: #e5e7eb; }
            .ganttTickLabel { position: absolute; top: 8px; left: 6px; font-size: 11px; color: #6b7280; white-space: nowrap; }
            .ganttLane { position: relative; height: 28px; border: 1px solid #e5e7eb; border-radius: 10px; background: #fff; overflow: hidden; }
            .ganttBar { position: absolute; top: 6px; height: 16px; border-radius: 999px; background: #111827; }
            .ganttBarMissing { position: absolute; inset: 0; display: flex; align-items: center; justify-content: center; font-size: 11px; color: #9ca3af; }
            .pill { display: inline-flex; align-items: center; padding: 5px 8px; border-radius: 999px; border: 1px solid #e5e7eb; font-size: 12px; white-space: nowrap; }
            .muted { color: #6b7280; font-size: 12px; }
            @media print { button { display: none !important; } }
          </style>
        </head>
        <body>
          <div class="wrap">
            <h1>Programme Summary</h1>
            <div class="sub">Generated: ${escapeHtml(isoToday())}</div>
            ${html}
          </div>
          <script>window.onload = () => window.print();</script>
        </body>
      </html>
    `);
    w.document.close();
  }

  /* ------------ build chat context ------------ */
const YESNO = (b) => (b ? "YES" : "NO");

const chatContext = useMemo(() => {
  const overdue = summaryItems.filter((x) => x.status === "overdue").slice(0, 30);

  const dueSoon = summaryItems
    .filter((x) => x.status !== "done" && x.statusA)
    .slice()
    .sort((a, b) => (parseISO(a.statusA)?.getTime() ?? 9e15) - (parseISO(b.statusA)?.getTime() ?? 9e15))
    .slice(0, 30);

  const bySupplier = {};
  summaryItems.forEach((x) => {
    const key = x.supplier || "—";
    bySupplier[key] = bySupplier[key] || { overdue: 0, ongoing: 0, done: 0, total: 0 };
    bySupplier[key].total += 1;
    bySupplier[key][x.status] += 1;
  });

  // --- programme helpers (inside useMemo so scope is guaranteed) ---
  const parseDateUTC = (s) => {
  if (!s) return null;
  const str = String(s).trim();

  // ISO: YYYY-MM-DD
  let m = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    const y = Number(m[1]), mo = Number(m[2]), d = Number(m[3]);
    const dt = new Date(Date.UTC(y, mo - 1, d));
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  // UK: DD/MM/YYYY
  m = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) {
    const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
    const dt = new Date(Date.UTC(y, mo - 1, d));
    return Number.isNaN(dt.getTime()) ? null : dt;
  }

  return null;
};

  const monthKeyUTC = (dateStr) => {
  const d = parseDateUTC(dateStr);
  if (!d) return "";
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${y}-${m}`;
};

const monthNameUTC = (dateStr) => {
  const d = parseDateUTC(dateStr);
  if (!d) return "";
  return d.toLocaleString("en-GB", { month: "long", timeZone: "UTC" });
};

const monthsBetweenUTC = (startStr, endStr) => {
  const s = parseDateUTC(startStr);
  const e = parseDateUTC(endStr);
  if (!s || !e) return [];

  const start = new Date(Date.UTC(s.getUTCFullYear(), s.getUTCMonth(), 1));
  const end = new Date(Date.UTC(e.getUTCFullYear(), e.getUTCMonth(), 1));

  const out = [];
  let cur = start;

  while (cur.getTime() <= end.getTime()) {
    const y = cur.getUTCFullYear();
    const m = String(cur.getUTCMonth() + 1).padStart(2, "0");
    out.push(`${y}-${m}`);
    cur = new Date(Date.UTC(y, cur.getUTCMonth() + 1, 1));
  }
  return out;
};

  const monthNameFromKeyUTC = (ym) => {
    const match = String(ym || "").match(/^(\d{4})-(\d{2})$/);
    if (!match) return "";
    const y = Number(match[1]);
    const mo = Number(match[2]);
    const d = new Date(Date.UTC(y, mo - 1, 1));
    return d.toLocaleString("en-GB", { month: "long", timeZone: "UTC" });
  };

  // Programme index (master) so chat can answer start/finish/duration + on-site (range)
const programme = (projects || []).flatMap((p) =>
  (p.master || []).flatMap((m) =>
    (m.levels || []).map((lv) => {
      const startKey = monthKeyUTC(lv.startDate);
      const finishKey = monthKeyUTC(lv.finishDate);

      const onSiteMonths = monthsBetweenUTC(lv.startDate, lv.finishDate);
      const onSiteMonthNames = onSiteMonths.map(monthNameFromKeyUTC);

      return {
        project: p.name || "",
        blockZone: m.blockZone || "",
        level: lv.name || "",

        startDate: lv.startDate || "",
        finishDate: lv.finishDate || "",
        durationDays: lv.startDate && lv.finishDate ? diffDaysUTC(lv.startDate, lv.finishDate) : null,

        startMonthKey: startKey,
        finishMonthKey: finishKey,
        startMonthName: monthNameUTC(lv.startDate),
        finishMonthName: monthNameUTC(lv.finishDate),

        onSiteRange: lv.startDate && lv.finishDate ? `${lv.startDate} → ${lv.finishDate}` : "",
        onSiteMonths,
        onSiteMonthNames,

        searchText: [
          p.name,
          m.blockZone,
          lv.name,
          `start:${lv.startDate || ""}`,
          `finish:${lv.finishDate || ""}`,
          startKey && `startMonth:${startKey}`,
          finishKey && `finishMonth:${finishKey}`,
          onSiteMonths.length ? `onSiteMonths:${onSiteMonths.join(",")}` : "",
          onSiteMonthNames.length ? `onSiteMonthNames:${onSiteMonthNames.join(",")}` : "",
          lv.startDate && lv.finishDate ? `onSite:${lv.startDate}..${lv.finishDate}` : "",
        ]
          .filter(Boolean)
          .join(" | "),
      };
    })
  )
);
  
// ✅ NEW: month -> projects/stages index (ensures chat returns ALL projects)
const programmeIndex = {};
programme.forEach((s) => {
  const months = Array.isArray(s.onSiteMonths) ? s.onSiteMonths : [];
  months.forEach((mk) => {
    if (!programmeIndex[mk]) programmeIndex[mk] = { projects: new Set(), stages: [] };
    programmeIndex[mk].projects.add(s.project);
    programmeIndex[mk].stages.push({
      project: s.project,
      blockZone: s.blockZone,
      level: s.level,
      startDate: s.startDate,
      finishDate: s.finishDate,
      onSiteRange: s.onSiteRange,
    });
  });
});

// JSON-serialisable (Set -> Array)
Object.keys(programmeIndex).forEach((mk) => {
  programmeIndex[mk].projects = Array.from(programmeIndex[mk].projects).sort((a, b) => a.localeCompare(b));
});

return {
  app: "SupplySync",
  today: isoToday(),
  view,

  activeProject: activeProject ? { id: activeProject.id, name: activeProject.name } : null,
  activePage: activePage ? { id: activePage.id, name: activePage.name, isMaster: !!activePage.meta?.isMaster } : null,

  counts: {
    projects: projects.length,
    summaryItems: summaryItems.length,
    overdue: summaryItems.filter((x) => x.status === "overdue").length,
    ongoing: summaryItems.filter((x) => x.status === "ongoing").length,
    done: summaryItems.filter((x) => x.status === "done").length,
  },

  sample: {
    overdueTop: overdue.map((x) => ({
      project: x.projectName,
      responsibility: x.pageName,
      supplier: x.supplier || "—",
      item: x.title,
      requiredOnSite: x.requiredOnSite,
      statusA: x.statusA,
      traffic: x.traffic,
      ticks: {
        statusA: YESNO(!!x.statusADone),
        firstIssue: YESNO(!!x.firstIssueDone),
        completed: YESNO(!!x.completed),
        notRequired: YESNO(!!x.notRequired),
      },
    })),

    upcomingStatusA: dueSoon.map((x) => ({
      project: x.projectName,
      responsibility: x.pageName,
      supplier: x.supplier || "—",
      item: x.title,
      statusA: x.statusA,
      traffic: x.traffic,
      ticks: {
        statusA: YESNO(!!x.statusADone),
        firstIssue: YESNO(!!x.firstIssueDone),
        completed: YESNO(!!x.completed),
        notRequired: YESNO(!!x.notRequired),
      },
    })),
  },

  bySupplier,

  // searchable programme data (now supports on-site month spans)
  programme,
  programmeIndex,

  // Full searchable tracker index for the assistant
  rows: summaryItems.map((x) => {
    const projectId = x.projectId ?? "";
    const pageId = x.pageId ?? "";
    const rowId = x.rowId ?? "";

    const key = `${projectId}:${pageId}:${rowId}`;

    const ticksBool = {
      statusA: !!x.statusADone,
      firstIssue: !!x.firstIssueDone,
      done: !!x.completed,
      notRequired: !!x.notRequired,
    };

    const ticks = {
      statusA: YESNO(ticksBool.statusA),
      firstIssue: YESNO(ticksBool.firstIssue),
      completed: YESNO(ticksBool.done),
      notRequired: YESNO(ticksBool.notRequired),
    };

    const commentsAll = x.commentsText || "";
    const comments = commentsAll.length > 900 ? commentsAll.slice(0, 900) + "…" : commentsAll;

    const searchText = [
      x.projectName,
      x.pageName,
      x.supplier || "—",
      x.title,
      x.requiredOnSite,
      x.statusA,
      x.firstIssue,
      x.timeframe,
      `totalDurationDays:${x.totalDurationDays ?? ""}`,
      `status:${x.status}`,
      `traffic:${x.traffic}`,
      `ticks:${Object.entries(ticks)
        .map(([k, v]) => `${k}=${v}`)
        .join(",")}`,
      comments,
    ]
      .filter(Boolean)
      .join(" | ");

    return {
      key,
      projectId,
      pageId,
      rowId,

      project: x.projectName,
      responsibility: x.pageName,
      supplier: x.supplier || "—",
      item: x.title,

      requiredOnSite: x.requiredOnSite,
      statusA: x.statusA,
      firstIssue: x.firstIssue,

      timeframe: x.timeframe,
      daysReqToStatusA: x.daysReqToStatusA,
      daysStatusAToFirstIssue: x.daysStatusAToFirstIssue,
      totalDurationDays: x.totalDurationDays,

      ticksBool,
      ticks,

      comments,
      commentsCount: x.commentsCount || 0,

      status: x.status,
      traffic: x.traffic,

      searchText,
    };
  }),
};

}, [summaryItems, view, activeProject, activePage, projects]);



  async function sendChat() {
    const text = chatInput.trim();
    if (!text || chatBusy) return;

    const next = [...chatMessages, { role: "user", content: text }];
    setChatMessages(next);
    setChatInput("");
    setChatBusy(true);
    setChatStatus("checking");

    try {
      const res = await fetch("/api/chat", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messages: next, context: chatContext, searchBackups: !!chatSearchBackups }),
      });

      if (!res.ok) {
        setChatStatus("error");
        setChatMessages((m) => [...m, { role: "assistant", content: `Chat error: server returned ${res.status}.` }]);
        return;
      }

      const data = await res.json().catch(() => ({}));
      const answer = data?.answer || data?.message || "No response returned from server.";
      setChatMessages((m) => [...m, { role: "assistant", content: answer }]);
      setChatStatus("ok");
    } catch {
      setChatStatus("error");
      setChatMessages((m) => [
        ...m,
        { role: "assistant", content: "Chat error: could not reach /api/chat. Check your server is running and route exists." },
      ]);
    } finally {
      setChatBusy(false);
    }
  }

  /* ---------- boot routing + loading screen ---------- */
  const isBooting = isHydrating || authUser === undefined;
  const didInitialRouteRef = useRef(false);

  useEffect(() => {
    if (isBooting) return;
    if (didInitialRouteRef.current) return;

    // Only decide initial landing once we have BOTH auth + blob state
    didInitialRouteRef.current = true;

    if (isAdmin) {
      // Admin always starts on Home
      setView(VIEW.LANDING);
      return;
    }

    // Guests land on their first allowed tracker page
    const proj = (visibleProjects || [])[0] || null;
    if (!proj) return;

    setActiveProjectId(proj.id);

    const allowed = (proj.pages || []).filter((pg) => {
      if (pg.meta?.isMaster) return false;
      const respId = pg.meta?.responsibilityId;
      const resp = (proj.responsibilities || []).find((r) => r.id === respId);
      const supplier = String(resp?.supplier || "").trim();
      return supplier && hasSupplierAccess(supplier);
    });

    if (allowed.length) setActivePageId(allowed[0].id);
    setView(VIEW.PROJECT);
  }, [isBooting, isAdmin, visibleProjects, hasSupplierAccess]);

  if (isBooting) {
    return (
      <div style={{ minHeight: "100vh", background: "#F9FAFB" }}>
        <div style={styles.loadingTopBar} />
        <div style={{ padding: 24, maxWidth: 720, margin: "0 auto" }}>
          <div style={styles.loadingCard}>
            <div style={styles.loadingTitle}>Loading latest tracker data…</div>
            <div style={styles.loadingSub}>
              Syncing the most recent data.
            </div>
            {hydrateError ? (
              <div style={{ marginTop: 10, fontSize: 12, color: "#B91C1C" }}>{hydrateError}</div>
            ) : null}
          </div>
        </div>
      </div>
    );
  }


  const saveButton = (
    <button
      onClick={() => {
        // Force an immediate save (clears debounce)
        clearTimeout(saveTimerRef.current);
        const payload = {
          globalDaysReqToStatusA,
          globalDaysStatusAToFirstIssue,
          projects,
          activeProjectId,
          activePageId,
          view,
          summaryFilter,
          summaryProjectId,
          summarySupplier,
        };
        saveToBlobNow(payload);
      }}
      title={saveStatus === "error" ? saveError || "Save failed" : "Saved to Azure Blob"}
      style={{
        ...styles.savePill,
        background:
          saveStatus === "saving" ? "#FEF3C7" : saveStatus === "saved" ? "#D1FAE5" : saveStatus === "error" ? "#FEE2E2" : "#F3F4F6",
        color:
          saveStatus === "saving" ? "#92400E" : saveStatus === "saved" ? "#065F46" : saveStatus === "error" ? "#991B1B" : "#111827",
      }}
    >
      {saveStatus === "saving"
        ? "Saving…"
        : saveStatus === "saved"
        ? `Saved${lastSavedAt ? ` • ${new Date(lastSavedAt).toLocaleTimeString()}` : ""}`
        : saveStatus === "error"
        ? "Save failed"
        : "Saved"}
    </button>
  );

  /* ---------- VIEW: LANDING ---------- */
  if (view === VIEW.LANDING) {
    return (
      <>
        <div style={styles.shell}>
          <div style={styles.page}>
            <div style={styles.header}>
              <div style={styles.brandRow}>
                <img src="/supplysync-logo.PNG" alt="SupplySync" style={styles.brandLogo} />
                <div>
                  <h1 style={styles.h1}>SupplySync</h1>
                  <p style={styles.sub}>Your Strategic Supply & Delivery Platform</p>
                </div>
              </div>
              <div style={styles.headerButtons}>{saveButton}</div>
            </div>

            <div style={styles.card}>
              <div style={styles.sectionTitle}>Project</div>
              <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap", marginTop: 8 }}>
        <select
          style={{ ...styles.input, width: 260 }}
          value={activeProject?.id || ""}
          onChange={(e) => {
            const pid = e.target.value;
            setActiveProjectId(pid);
            const proj = (visibleProjects || []).find((p) => p.id === pid) || projects.find((p) => p.id === pid);

            if (!proj) {
              setActivePageId(null);
              return;
            }

            if (isGuest) {
              const allowed = (proj.pages || []).filter((pg) => {
                if (pg.meta?.isMaster) return false;
                const respId = pg.meta?.responsibilityId;
                const resp = (proj.responsibilities || []).find((r) => r.id === respId);
                const supplier = String(resp?.supplier || "").trim();
                return supplier && hasSupplierAccess(supplier);
              });
              setActivePageId(allowed[0]?.id || null);
            } else {
              const mp = proj?.pages?.find((x) => x.meta?.isMaster) || proj?.pages?.[0];
              setActivePageId(mp?.id || null);
            }
          }}
        >
          {(visibleProjects || projects).map((p) => (
                    <option key={p.id} value={p.id}>
                      {p.name}
                    </option>
                  ))}
                </select>

                <button style={styles.primaryBtn} onClick={addProject}>
                  + Project
                </button>

                <button style={styles.secondaryBtn} onClick={goToProject}>
                  Go to Project
                </button>
              
              <button
  style={styles.dangerBtn}
  onClick={() => {
    if (!activeProject) return;

    const ok = window.confirm(
      `Are you sure you want to remove "${activeProject.name}"?\n\n` +
      `This cannot be undone.\n\n` +
      `If a backup exists from before the last save, the project can be recovered from backup.`
    );    

    if (!ok) return;

    setProjects((prev) => prev.filter((p) => p.id !== activeProject.id));
    setActiveProjectId(null);
    setActivePageId(null);
    setView(VIEW.LANDING);
  }}
>
  Remove Project
</button>
                </div>

              <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginTop: 14 }}>
                <button
                  style={styles.secondaryBtn}
                  onClick={() => {
                    setSummaryProjectId("all");
                    setSummarySupplier("all");
                    setView(VIEW.SUMMARY);
                  }}
                >
                  Summary
                </button>
                <button style={styles.secondaryBtn} onClick={() => setView(VIEW.GANTT_SUMMARY)}>
                  Programme Summary (Gantt)
                </button>
              </div>
            </div>
          </div>
        </div>

        {!isGuest && (
          <ChatOverlay
            open={chatOpen}
            setOpen={setChatOpen}
            messages={chatMessages}
            input={chatInput}
            setInput={setChatInput}
            searchBackups={chatSearchBackups}
            setSearchBackups={setChatSearchBackups}
            busy={chatBusy}
            sendChat={sendChat}
            endRef={chatEndRef}
            status={chatStatus}
            onReset={resetChat}
          />
        )}
      </>
    );
  }

  /* ---------- VIEW: PROGRAMME SUMMARY (FULL SCREEN) ---------- */
  if (view === VIEW.GANTT_SUMMARY) {
    return (
      <>
        <div style={styles.fullscreen}>
          <div style={styles.fullTopBar}>
            <div>
              <div style={{ fontWeight: 900, fontSize: 16 }}>Programme Summary</div>
              <div style={styles.muted}>All projects, stacked. Print will output this page only.</div>
            </div>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button style={styles.secondaryBtn} onClick={() => setView(VIEW.LANDING)}>
                Home
              </button>
              <button style={styles.secondaryBtn} onClick={() => setView(VIEW.SUMMARY)}>
                Summary
              </button>
              <button style={styles.primaryBtn} onClick={printProgrammeSummary}>
                Print
              </button>
            </div>
          </div>

          <div style={styles.fullBody}>
            <div style={styles.fullBodyInner} ref={programmePrintRef}>
              {projects.map((p) => (
                <div key={p.id} style={styles.projectSection}>
                  <div style={styles.projectHeader}>
                    <div>
                      <div style={styles.projectTitle}>{p.name}</div>
                      <div style={styles.muted}>Programme from Block/Zone + Level start/finish dates</div>
                    </div>
                    <button
                      style={styles.secondaryBtn}
                      onClick={() => {
                        setActiveProjectId(p.id);
                        const mp = p.pages.find((x) => x.meta?.isMaster) || p.pages[0];
                        setActivePageId(mp?.id || null);
                        setView(VIEW.PROJECT);
                      }}
                    >
                      Open Project
                    </button>
                  </div>

                  <div style={styles.projectGantt}>
                    <ProgrammeGantt master={p.master || []} dense />
                  </div>
                </div>
              ))}
            </div>
          </div>
        </div>

        {!isGuest && (
          <ChatOverlay
            open={chatOpen}
            setOpen={setChatOpen}
            messages={chatMessages}
            input={chatInput}
            setInput={setChatInput}
            searchBackups={chatSearchBackups}
            setSearchBackups={setChatSearchBackups}
            busy={chatBusy}
            sendChat={sendChat}
            endRef={chatEndRef}
            status={chatStatus}
            onReset={resetChat}
          />
        )}
      </>
    );
  }

  /* ---------- VIEW: SUMMARY ---------- */
  if (view === VIEW.SUMMARY) {
    return (
      <>
        <div style={styles.shell}>
          <div style={styles.page}>
            <div style={styles.header}>
              <div style={styles.brandRow}>
                <img src="/supplysync-logo.PNG" alt="SupplySync" style={styles.brandLogo} />
                <div>
                  <h1 style={styles.h1}>Summary</h1>
                  <p style={styles.sub}>All projects + responsibilities (excluding “Not required”). Traffic is based on Status A.</p>
                </div>
              </div>
              <div style={styles.headerButtons}>
                {saveButton}
                <button style={styles.secondaryBtn} onClick={() => setView(VIEW.LANDING)}>
                  Home
                </button>
                <button style={styles.secondaryBtn} onClick={() => setView(VIEW.GANTT_SUMMARY)}>
                  Programme Summary
                </button>
                <button style={styles.secondaryBtn} onClick={exportSummaryCSV}>
                  Export CSV
                </button>
                <button style={styles.secondaryBtn} onClick={exportSummaryPDF}>
                  Export PDF
                </button>
              </div>
            </div>

            <div style={styles.card}>
              <div style={styles.tableTop}>
                <div style={{ display: "grid", gap: 8 }}>
                  <TrafficKeyDotsOnly />

                  <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
                    <label style={{ display: "grid", gap: 4, fontSize: 12, color: "#374151" }}>
                      Project
                      <select style={{ ...styles.input, width: 220 }} value={summaryProjectId} onChange={(e) => setSummaryProjectId(e.target.value)}>
                        <option value="all">All projects</option>
                        {projects.map((p) => (
                          <option key={p.id} value={p.id}>
                            {p.name}
                          </option>
                        ))}
                      </select>
                    </label>

                    <label style={{ display: "grid", gap: 4, fontSize: 12, color: "#374151" }}>
                      Supplier
                      <select style={{ ...styles.input, width: 220 }} value={summarySupplier} onChange={(e) => setSummarySupplier(e.target.value)}>
                        <option value="all">All suppliers</option>
                        {supplierOptions.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </label>
                  </div>
                </div>

                
              <div style={styles.tableWrap}>
                <table style={styles.table}>
                  <thead>
                    <tr>
                      <th style={styles.th}>●</th>
                      <th style={styles.th}>Project</th>
                      <th style={styles.th}>Responsibility</th>
                      <th style={styles.th}>Supplier</th>
                      <th style={styles.th}>Item</th>
                      <th style={styles.th}>Req</th>
                      <th style={styles.th}>Status A</th>
                      <th style={styles.th}>First</th>
                      <th style={styles.th}></th>
                    </tr>
                  </thead>
                  <tbody>
  {filteredSummary.map((it) => {
    const stripe = summaryBorderColor(it);

    return (
      <tr
        key={`${it.projectId}:${it.pageId}:${it.rowId}`}
        style={it.status === "overdue" ? styles.trLate : it.status === "done" ? styles.trDone : undefined}
      >
        
        <td style={styles.tdCenter}>
          <TrafficDot status={it.traffic} />
        </td>
        <td style={styles.td}>{it.projectName}</td>
        <td style={styles.td}>{it.pageName}</td>
        <td style={styles.td}>{it.supplier || "—"}</td>
        <td style={{ ...styles.td, ...styles.wrap }}>{it.title}</td>
        <td style={styles.td}>
  <div style={{ ...styles.pillCompact, ...datePillStyle({ row: it, dateKey: "requiredOnSite" }) }}>
    {it.requiredOnSite || "—"}
  </div>
</td>

<td style={styles.td}>
  <div style={{ ...styles.pillCompact, ...datePillStyle({ row: it, dateKey: "statusA" }) }}>
    {it.statusA || "—"}
  </div>
</td>

<td style={styles.td}>
  <div style={{ ...styles.pillCompact, ...datePillStyle({ row: it, dateKey: "firstIssue" }) }}>
    {it.firstIssue || "—"}
  </div>
</td>

        <td style={styles.td}>
          <button style={styles.smallBtn} onClick={() => jumpToItem(it)}>
            Open
          </button>
        </td>
      </tr>
    );
  })}

  {!filteredSummary.length ? (
    <tr>
      <td style={styles.td} colSpan={10}>
        <div style={{ color: "#6B7280" }}>No items.</div>
      </td>
    </tr>
  ) : null}
</tbody>


                </table>
              </div>
            </div>
          </div>
        </div>

        {!isGuest && (
          <ChatOverlay
            open={chatOpen}
            setOpen={setChatOpen}
            messages={chatMessages}
            input={chatInput}
            setInput={setChatInput}
            searchBackups={chatSearchBackups}
            setSearchBackups={setChatSearchBackups}
            busy={chatBusy}
            sendChat={sendChat}
            endRef={chatEndRef}
            status={chatStatus}
            onReset={resetChat}
          />
        )}
      </>
    );
  }

  /* ---------- VIEW: PROJECT ---------- */
  return (
    <>
      <div style={styles.shell}>
        <ProjectView
          projects={projects}
          visibleProjects={visibleProjects}
          activeProjectId={activeProjectId}
          activePageId={activePageId}
          activeProject={activeProject}
          activePage={activePage}
          authUser={authUser}
          isAdmin={isAdmin}
          isGuest={isGuest}
          hasSupplierAccess={hasSupplierAccess}
          setActiveProjectId={setActiveProjectId}
          setActivePageId={setActivePageId}
          updateProject={updateProject}
          updatePage={updatePage}
          updateRow={updateRow}
          computedRows={computedRows}
          addProject={addProject}
          addMasterRow={addMasterRow}
          updateMasterRow={updateMasterRow}
          removeMasterRow={removeMasterRow}
          addLevel={addLevel}
          updateLevel={updateLevel}
          removeLevel={removeLevel}
          addResponsibility={addResponsibility}
          updateResponsibility={updateResponsibility}
          removeResponsibility={removeResponsibility}
          addManualRow={addManualRow}
          removeRow={removeRow}
          globalDaysReqToStatusA={globalDaysReqToStatusA}
          globalDaysStatusAToFirstIssue={globalDaysStatusAToFirstIssue}
          setGlobalDaysReqToStatusA={setGlobalDaysReqToStatusA}
          setGlobalDaysStatusAToFirstIssue={setGlobalDaysStatusAToFirstIssue}
          isMasterPage={isMasterPage}
          setView={setView}
          VIEW={VIEW}
          saveButton={saveButton}
        />
      </div>

      {!isGuest && (
        <ChatOverlay
          open={chatOpen}
          setOpen={setChatOpen}
          messages={chatMessages}
          input={chatInput}
          setInput={setChatInput}
          searchBackups={chatSearchBackups}
          setSearchBackups={setChatSearchBackups}
          busy={chatBusy}
          sendChat={sendChat}
          endRef={chatEndRef}
          status={chatStatus}
          onReset={resetChat}
        />
      )}
    </>
  );
}

/* ---------- Project View ---------- */
function ProjectView(props) {
  const {
    projects,
    visibleProjects,
    activeProjectId,
    activePageId,
    activeProject,
    activePage,
    authUser,
    isAdmin,
    isGuest,
    hasSupplierAccess,
    setActiveProjectId,
    setActivePageId,
    updateProject,
    updateRow,
    computedRows,
    addProject,
    addMasterRow,
    updateMasterRow,
    removeMasterRow,
    addLevel,
    updateLevel,
    removeLevel,
    addResponsibility,
    updateResponsibility,
    removeResponsibility,
    addManualRow,
    removeRow,
    globalDaysReqToStatusA,
    globalDaysStatusAToFirstIssue,
    setGlobalDaysReqToStatusA,
    setGlobalDaysStatusAToFirstIssue,
    isMasterPage,
    setView,
    VIEW,
    saveButton,
  } = props;

  // Pages a guest is allowed to see inside the active project
  const allowedPages = useMemo(() => {
    if (!activeProject?.pages?.length) return [];
    if (isAdmin) return activeProject.pages;
    return (activeProject.pages || []).filter((pg) => {
      if (pg.meta?.isMaster) return false;
      const respId = pg.meta?.responsibilityId;
      const resp = (activeProject.responsibilities || []).find((r) => r.id === respId);
      const supplier = String(resp?.supplier || "").trim();
      return supplier && hasSupplierAccess(supplier);
    });
  }, [activeProject, isAdmin, hasSupplierAccess]);

  // If a guest lands on Project Home, immediately jump to their first allowed tracker page
  useEffect(() => {
    if (!isGuest) return;
    if (!activeProject || !activePage) return;
    if (!activePage.meta?.isMaster) return;
    if (!allowedPages.length) return;
    setActivePageId(allowedPages[0].id);
  }, [isGuest, activeProject, activePage, allowedPages, setActivePageId]);

  // Comments modal state (append-only; each saved comment is locked)
  const [commentRowId, setCommentRowId] = useState(null);
  const [commentName, setCommentName] = useState("");
  const [commentText, setCommentText] = useState("");

  const commentRow = useMemo(() => {
    if (!commentRowId) return null;
    const rows = activePage?.rows || [];
    return rows.find((r) => r.id === commentRowId) || null;
  }, [commentRowId, activePage]);

  function openComments(row) {
    if (!row || row.kind === "header") return;
    setCommentRowId(row.id);
    setCommentName(authUser?.userDetails || "");
    setCommentText("");
  }

  function closeComments() {
    setCommentRowId(null);
    setCommentText("");
  }

  function saveComment() {
    if (!commentRow) return;
    const name = clean(commentName) || (authUser?.userDetails || authUser?.userId || "guest");
    const text = clean(commentText);
    if (!text) return;

    const next = {
      id: uid(),
      name,
      dateISO: isoToday(),
      text,
      locked: true,
      lockedBy: authUser?.userDetails || authUser?.userId || "guest",
      lockedAt: new Date().toISOString(),
    };

    updateRow(commentRow.id, {
      comments: [...(Array.isArray(commentRow.comments) ? commentRow.comments : []), next],
    });

    setCommentText("");
  }

  // Guest: if they land on Project Home, push them to their first allowed tracker page
  useEffect(() => {
    if (!activeProject || !activePage) return;
    if (!isGuest) return;
    if (authUser === undefined) return;
    if (!activePage.meta?.isMaster) return;

    const first = allowedPages[0] || null;
    if (first && first.id !== activePage.id) setActivePageId(first.id);
  }, [activeProject, activePage, allowedPages, isGuest, authUser, setActivePageId]);

 // ---------------- Locks: guest confirm -> lock, admin unlock ----------------
function lockInfo() {
  return {
    lockedBy: authUser?.userDetails || authUser?.userId || "guest",
    lockedAt: new Date().toISOString(),
  };
}

function adminUnlock(rowId, field) {
  const row = activePage?.rows?.find((x) => x.id === rowId);
  const nextLocks = { ...(row?.locks || {}), [field]: null };
  updateRow(rowId, { locks: nextLocks });
}

function tickMilestone(row, field, checked) {
  if (!row) return;

  const locked = !!row?.locks?.[field];

  // Guests: can only tick ONCE if not locked; cannot untick
  if (isGuest) {
    if (!checked) return;          // guests cannot untick
    if (locked) return;            // locked = can't change

    const ok = window.confirm("Once ticked this will be locked and only administrator can unlock");
    if (!ok) return;

    updateRow(row.id, {
      [field]: true,
      locks: { ...(row.locks || {}), [field]: lockInfo() },
    });
    return;
  }

  // Admin: can tick/untick. Unticking clears the lock so guests can tick again.
  if (!checked) {
    updateRow(row.id, {
      [field]: false,
      locks: { ...(row.locks || {}), [field]: null }, // ✅ clears padlock
    });
    return;
  }

  // Admin ticking ON: keep existing lock meta if any; otherwise leave unlocked
  updateRow(row.id, {
    [field]: true,
    locks: { ...(row.locks || {}), [field]: row.locks?.[field] || null },
  });
}
  // ✅ selector block in same place for BOTH Project Home and responsibility pages
  const projectOptions = isAdmin ? projects : visibleProjects?.length ? visibleProjects : projects;

  const SelectorBar = () => (
    <div style={styles.selectorBar}>
      <div style={styles.selectorBlock}>
        <div style={styles.sectionTitle}>Project</div>
        <select
          style={{ ...styles.input, width: 260 }}
          value={activeProjectId || projectOptions?.[0]?.id || ""}
          onChange={(e) => {
            const pid = e.target.value;
            setActiveProjectId(pid);
            const proj = (visibleProjects || projects).find((p) => p.id === pid) || null;
            if (!proj) return;
            if (isAdmin) {
              const mp = proj.pages?.find((x) => x.meta?.isMaster) || proj.pages?.[0];
              setActivePageId(mp?.id || null);
            } else {
              const firstAllowed = (proj.pages || [])
                .filter((pg) => !pg.meta?.isMaster)
                .filter((pg) => {
                  const respId = pg.meta?.responsibilityId;
                  const resp = (proj.responsibilities || []).find((r) => r.id === respId);
                  const supplier = String(resp?.supplier || "").trim();
                  return supplier && hasSupplierAccess(supplier);
                })[0];
              setActivePageId(firstAllowed?.id || null);
            }
          }}
        >
          {(projectOptions || []).map((p) => (
            <option key={p.id} value={p.id}>
              {p.name}
            </option>
          ))}
        </select>
      </div>

      <div style={styles.selectorBlock}>
        <div style={styles.sectionTitle}>Pages</div>
        <select
          style={{ ...styles.input, width: 280 }}
          value={activePageId || allowedPages?.[0]?.id || ""}
          onChange={(e) => setActivePageId(e.target.value)}
          disabled={!activeProject}
        >
          {(allowedPages || []).map((pg) => (
            <option key={pg.id} value={pg.id}>
              {pg.name}
              {pg.meta?.isMaster ? " (Project Home)" : pg.meta?.generated ? " (Auto)" : ""}
            </option>
          ))}
        </select>
      </div>
    </div>
  );

  return (
    <div style={styles.page}>
      <div style={styles.header}>
        <div style={styles.brandRow}>
          <img src="/supplysync-logo.PNG" alt="SupplySync" style={styles.brandLogo} />
          <div>
            <h1 style={styles.h1}>{activeProject?.name || "Project"}</h1>
            <p style={styles.sub}>Project Home defines Blocks/Zones + Levels. Tracker pages auto-populate. Traffic is based on Status A.</p>
          </div>
        </div>
        <div style={styles.headerButtons}>
                {saveButton}
          {isAdmin ? (
            <>
              <button style={styles.secondaryBtn} onClick={() => setView(VIEW.LANDING)}>
                Home
              </button>
              <button style={styles.secondaryBtn} onClick={() => setView(VIEW.SUMMARY)}>
                Summary
              </button>
              <button style={styles.secondaryBtn} onClick={() => setView(VIEW.GANTT_SUMMARY)}>
                Programme Summary
              </button>
            </>
          ) : (
            <div style={styles.muted}>Guest access</div>
          )}
        </div>
      </div>

      {/* ✅ always left aligned, and ✅ NO "+ Project" here */}
      <div style={styles.card}>
        <SelectorBar />
        <div style={{ marginTop: 10 }}>
          <input
            style={{ ...styles.input, width: "100%" }}
            value={activeProject?.name || ""}
            onChange={(e) => activeProject && updateProject(activeProject.id, { name: e.target.value })}
            placeholder="Project name"
            disabled={!isAdmin}
          />
        </div>
      </div>

     {/* Defaults (hidden on Project Home + hidden for guests) */}
{!isMasterPage && !isGuest ? (
  <div style={styles.card}>
    <div style={styles.cardHeader}>
      <h2 style={styles.h2}>Default timeframes</h2>
      <span style={styles.muted}>(days between stages – can be overridden per row)</span>
    </div>
    <div style={styles.grid2}>
      <label style={styles.label}>
        Days: Req → Status A
        <input
          style={styles.input}
          type="number"
          min={0}
          value={globalDaysReqToStatusA}
          onChange={(e) => setGlobalDaysReqToStatusA(clampInt(e.target.value, 0))}
        />
      </label>
      <label style={styles.label}>
        Days: Status A → First
        <input
          style={styles.input}
          type="number"
          min={0}
          value={globalDaysStatusAToFirstIssue}
          onChange={(e) => setGlobalDaysStatusAToFirstIssue(clampInt(e.target.value, 0))}
        />
      </label>
    </div>
  </div>
) : null}


      {/* Project Home */}
      {isMasterPage ? (
        <div style={styles.card}>
          <div style={styles.tableTop}>
            <div>
              <h2 style={styles.h2}>Project Home</h2>
              <div style={styles.muted}>Start date populates “Required on Site” on tracker pages.</div>
            </div>
          </div>

          <div style={{ marginTop: 10 }}>
            <div style={styles.tableTop}>
              <div>
                <h3 style={styles.h3}>Blocks / Zones</h3>
              </div>
              <button style={styles.primaryBtn} onClick={addMasterRow} disabled={!activeProject}>
                + Block / Zone
              </button>
            </div>

            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Block / Zone</th>
                    <th style={styles.th}>Level</th>
                    <th style={styles.th}>Start</th>
                    <th style={styles.th}>Finish</th>
                    <th style={styles.th}>Duration</th>
                    <th style={styles.th}></th>
                  </tr>
                </thead>
                <tbody>
                  {(activeProject?.master || []).map((m) => (
                    <React.Fragment key={m.id}>
                      {(m.levels || []).map((lv, idx) => {
                        const dur = diffDaysUTC(lv.startDate, lv.finishDate);
                        const safeDur = dur == null ? null : dur;
                        const barW = safeDur == null ? 0 : Math.max(6, Math.min(180, safeDur * 6));
                        const isNeg = safeDur != null && safeDur < 0;

                        return (
                          <tr key={lv.id}>
                            {idx === 0 ? (
                              <td style={styles.td} rowSpan={(m.levels || []).length}>
                                <input
                                  style={styles.input}
                                  placeholder="e.g. Block A / Zone 1"
                                  value={m.blockZone}
                                  onChange={(e) => updateMasterRow(m.id, { blockZone: e.target.value })}
                                />
                              </td>
                            ) : null}

                            <td style={styles.td}>
                              <input
                                style={styles.input}
                                placeholder={`Level ${idx + 1}`}
                                value={lv.name}
                                onChange={(e) => updateLevel(m.id, lv.id, { name: e.target.value })}
                              />
                            </td>

                            <td style={styles.td}>
                              <input
                                style={styles.input}
                                type="date"
                                value={lv.startDate}
                                onChange={(e) => updateLevel(m.id, lv.id, { startDate: e.target.value })}
                              />
                            </td>

                            <td style={styles.td}>
                              <input
                                style={styles.input}
                                type="date"
                                value={lv.finishDate}
                                onChange={(e) => updateLevel(m.id, lv.id, { finishDate: e.target.value })}
                              />
                            </td>

                            <td style={styles.td}>
                              <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                                <div
                                  style={{
                                    height: 10,
                                    width: 190,
                                    borderRadius: 999,
                                    border: "1px solid #E5E7EB",
                                    background: "#FFFFFF",
                                    overflow: "hidden",
                                  }}
                                  title={safeDur == null ? "Add start + finish to calculate duration" : `${safeDur} days`}
                                >
                                  <div
                                    style={{
                                      height: "100%",
                                      width: safeDur == null ? 0 : barW,
                                      background: isNeg ? "#EF4444" : "#0D9488",
                                      opacity: 0.85,
                                    }}
                                  />
                                </div>
                                <div style={{ ...styles.pillCompact, ...(safeDur == null ? styles.pillEmpty : null), ...(isNeg ? styles.pillLate : null) }}>
                                  {safeDur == null ? "—" : `${safeDur}d`}
                                </div>
                              </div>
                            </td>

                            <td style={styles.td}>
                              <div style={styles.inline}>
                                <button style={styles.smallBtn} onClick={() => addLevel(m.id)}>
                                  + Level
                                </button>
                                <button style={styles.iconBtn} onClick={() => removeLevel(m.id, lv.id)} title="Remove level">
                                  ✕
                                </button>
                                {idx === 0 ? (
                                  <button style={styles.iconBtn} onClick={() => removeMasterRow(m.id)} title="Remove block/zone">
                                    🗑
                                  </button>
                                ) : null}
                              </div>
                            </td>
                          </tr>
                        );
                      })}
                    </React.Fragment>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div style={{ marginTop: 14 }}>
            <div style={styles.tableTop}>
              <div>
                <h3 style={styles.h3}>Design Responsibilities</h3>
                <div style={styles.muted}>Each responsibility creates a page in this project.</div>
              </div>
              <button style={styles.primaryBtn} onClick={addResponsibility} disabled={!activeProject}>
                + Responsibility
              </button>
            </div>

            <div style={styles.tableWrap}>
              <table style={styles.table}>
                <thead>
                  <tr>
                    <th style={styles.th}>Responsibility (page name)</th>
                    <th style={styles.th}>Supplier</th>
                    <th style={styles.th}></th>
                  </tr>
                </thead>
                <tbody>
                  {(activeProject?.responsibilities || []).map((r) => (
                    <tr key={r.id}>
                      <td style={styles.td}>
                        <input
                          style={styles.input}
                          placeholder="e.g. MSA / NCCT / Stone"
                          value={r.name}
                          onChange={(e) => updateResponsibility(r.id, { name: e.target.value })}
                        />
                      </td>
                      <td style={styles.td}>
                        <input
                          style={styles.input}
                          placeholder="e.g. ABC Consultants"
                          value={r.supplier || ""}
                          onChange={(e) => updateResponsibility(r.id, { supplier: e.target.value })}
                        />
                      </td>
                      <td style={styles.td}>
                        <button style={styles.iconBtn} onClick={() => removeResponsibility(r.id)} title="Remove">
                          ✕
                        </button>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          <div style={{ marginTop: 14 }}>
            <div style={styles.tableTop}>
              <div>
                <h3 style={styles.h3}>Programme</h3>
                <div style={styles.muted}>Simple Gantt chart using Level start/finish dates.</div>
              </div>
            </div>
            <div style={styles.cardInset}>
              <ProgrammeGantt master={activeProject?.master || []} />
            </div>
          </div>
        </div>
      ) : (
        /* Tracker Page */
        <div style={styles.card}>
          <div style={styles.tableTop}>
            <div>
              <h2 style={styles.h2}>Design Tracker – {activePage?.name}</h2>
              <div style={styles.muted}>Traffic is based on Status A date.</div>
              <div style={{ marginTop: 8 }}>
                <TrafficKeyDotsOnly />
              </div>
            </div>
            {isAdmin ? (
              <button style={styles.primaryBtn} onClick={addManualRow} disabled={!activePage}>
                + Manual row
              </button>
            ) : null}
          </div>

          <div style={styles.tableWrap}>
            <table style={styles.table}>
              <thead>
                <tr>
                  <th style={styles.thSmall}>Done</th>
                  <th style={styles.thSmall}>NR</th>
                  <th style={styles.thSmall}>●</th>
                  <th style={styles.thWide}>Title</th>
                  <th style={styles.thMed}>From</th>
                  <th style={styles.thMed}>Anchor</th>
                  <th style={styles.thMed}>Req</th>
                  <th style={styles.thMed}>Status A</th>
                  <th style={styles.thMed}>First</th>
                      <th style={styles.thTf}>TF</th>
                      <th style={styles.thSmall}>💬</th>
                      <th style={styles.thSmall}></th>
                </tr>
              </thead>
              <tbody>
                {computedRows.map((r) => {
                  const rowStyle =
                    r.kind === "header"
                      ? styles.trHeader
                      : r.notRequired
                      ? styles.trNotRequired
                      : r._overdue?.overdue
                      ? styles.trLate
                      : r.completed
                      ? styles.trDone
                      : undefined;

                  const disabled = r.completed || r.notRequired;

                  return (
                    <tr key={r.id} style={rowStyle}>
                      <td style={styles.tdCenter}>
                        {r.kind === "header" ? null : (
                          <input
                            type="checkbox"
                            checked={!!r.completed}
                            onChange={(e) => updateRow(r.id, { completed: e.target.checked })}
                            disabled={r.notRequired || isGuest}
                          />
                        )}
                      </td>

                      <td style={styles.tdCenter}>
                        {r.kind === "header" ? null : (
                          <input
                            type="checkbox"
                            checked={!!r.notRequired}
                            onChange={(e) => {
                              const checked = e.target.checked;
                              updateRow(r.id, {
                                notRequired: checked,
                                ...(checked ? { completed: false, statusADone: false, firstIssueDone: false } : {}),
                              });
                            }}
                            disabled={isGuest}
                          />
                        )}
                      </td>

                      <td style={styles.tdCenter}>{r.kind === "header" ? null : <TrafficDot status={r._traffic} />}</td>

                      <td style={styles.td}>
                        {r.kind === "header" ? (
                          <div style={{ fontWeight: 800 }}>{r.item}</div>
                        ) : (
                          <input
                            style={{ ...styles.input, ...(r.notRequired ? styles.inputMuted : null) }}
                            value={r.item}
                            onChange={(e) => updateRow(r.id, { item: e.target.value, meta: { ...r.meta, generated: false } })}
                            disabled={r.notRequired || isGuest}
                          />
                        )}
                      </td>

                      <td style={styles.td}>
                        {r.kind === "header" ? null : (
                          <select
                            style={{ ...styles.input, ...(r.notRequired ? styles.inputMuted : null) }}
                            value={r.anchorKey}
                            onChange={(e) => updateRow(r.id, { anchorKey: e.target.value })}
                            disabled={r.notRequired || isGuest}
                          >
                            {ANCHORS.map((a) => (
                              <option key={a.key} value={a.key}>
                                {a.label}
                              </option>
                            ))}
                          </select>
                        )}
                      </td>

                      <td style={styles.td}>
                        {r.kind === "header" ? null : (
                          <input
                            style={{ ...styles.input, ...(r.notRequired ? styles.inputMuted : null) }}
                            type="date"
                            value={r.anchorDateISO}
                            onChange={(e) => updateRow(r.id, { anchorDateISO: e.target.value })}
                            disabled={r.notRequired || isGuest}
                          />
                        )}
                      </td>

                      <td style={styles.td}>
                        <DatePill value={r._computed.requiredOnSite} isHeader={r.kind === "header"} overdue={!!r._overdue?.overdueReq} done={r.completed} muted={r.notRequired} />
                      </td>

                      <td style={styles.td}>
                        <MilestoneCell
                          isHeader={r.kind === "header"}
                          value={r._computed.statusA}
                          checked={!!r.statusADone}
                          locked={!!r.locks?.statusADone}
                          lockMeta={r.locks?.statusADone}
                          isAdmin={isAdmin}
                          onUnlock={() => adminUnlock(r.id, "statusADone")}
                          onChange={(checked) => tickMilestone(r, "statusADone", checked)}
                          overdue={!!r._overdue?.overdueA}
                          disabled={disabled || (isGuest && !!r.locks?.statusADone)}
                          muted={r.notRequired}
                        />
                      </td>

                      <td style={styles.td}>
                        <MilestoneCell
                          isHeader={r.kind === "header"}
                          value={r._computed.firstIssue}
                          checked={!!r.firstIssueDone}
                          locked={!!r.locks?.firstIssueDone}
                          lockMeta={r.locks?.firstIssueDone}
                          isAdmin={isAdmin}
                          onUnlock={() => adminUnlock(r.id, "firstIssueDone")}
                          onChange={(checked) => tickMilestone(r, "firstIssueDone", checked)}
                          overdue={!!r._overdue?.overdueF}
                          disabled={disabled || (isGuest && !!r.locks?.firstIssueDone)}
                          muted={r.notRequired}
                        />
                      </td>

                      <td style={styles.td}>
                        {r.kind === "header" ? null : (
                          <div style={styles.tfRow}>
                            <input
                              style={{ ...styles.tfInput, ...(r.notRequired ? styles.inputMuted : null) }}
                              type="number"
                              min={0}
                              value={r.overrideDaysReqToStatusA ?? ""}
                              placeholder={String(globalDaysReqToStatusA)}
                              onChange={(e) => updateRow(r.id, { overrideDaysReqToStatusA: e.target.value === "" ? null : clampInt(e.target.value, 0) })}
                              disabled={r.notRequired || isGuest}
                              title="Req→A"
                            />
                            <input
                              style={{ ...styles.tfInput, ...(r.notRequired ? styles.inputMuted : null) }}
                              type="number"
                              min={0}
                              value={r.overrideDaysStatusAToFirstIssue ?? ""}
                              placeholder={String(globalDaysStatusAToFirstIssue)}
                              onChange={(e) => updateRow(r.id, { overrideDaysStatusAToFirstIssue: e.target.value === "" ? null : clampInt(e.target.value, 0) })}
                              disabled={r.notRequired || isGuest}
                              title="A→First"
                            />
                          </div>
                        )}
                      </td>

                      <td style={styles.tdCenter}>
                        {r.kind === "header" ? null : (
                          <button style={styles.smallBtn} onClick={() => openComments(r)} title="Add/view comments">
                            💬 {Array.isArray(r.comments) && r.comments.length ? r.comments.length : ""}
                          </button>
                        )}
                      </td>

                      <td style={styles.tdCenter}>
                        {r.kind === "header" ? null : isAdmin ? (
                          <button style={styles.iconBtn} onClick={() => removeRow(r.id)} title="Delete row">
                            ✕
                          </button>
                        ) : null}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {commentRowId ? (
        <div style={styles.modalOverlay} onMouseDown={closeComments}>
          <div style={styles.modalCard} onMouseDown={(e) => e.stopPropagation()}>
            <div style={styles.modalHeader}>
              <div style={{ fontWeight: 900 }}>Comments</div>
              <button style={styles.iconBtn} onClick={closeComments} title="Close">
                ✕
              </button>
            </div>

            <div style={{ display: "grid", gap: 10 }}>
              <div style={styles.muted}>Row: {commentRow?.item || "—"}</div>

              <label style={styles.label}>
                Your name
                <input style={styles.input} value={commentName} onChange={(e) => setCommentName(e.target.value)} placeholder="Name" />
              </label>

              <label style={styles.label}>
                Date
                <input style={{ ...styles.input, ...styles.inputMuted }} value={isoToday()} readOnly />
              </label>

              <label style={styles.label}>
                Comment
                <textarea
                  style={{ ...styles.input, minHeight: 90, resize: "vertical" }}
                  value={commentText}
                  onChange={(e) => setCommentText(e.target.value)}
                  placeholder="Write your comment…"
                />
              </label>

              <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
                <button style={styles.secondaryBtn} onClick={closeComments}>
                  Cancel
                </button>
                <button style={styles.primaryBtn} onClick={saveComment}>
                  Save (locks)
                </button>
              </div>

              <div style={{ borderTop: "1px solid #E5E7EB", paddingTop: 10 }}>
                <div style={{ fontWeight: 900, marginBottom: 6 }}>History</div>
                {Array.isArray(commentRow?.comments) && commentRow.comments.length ? (
                  <div style={{ display: "grid", gap: 8 }}>
                    {[...commentRow.comments]
                      .slice()
                      .reverse()
                      .map((c) => (
                        <div key={c.id} style={styles.commentItem}>
                          <div style={{ display: "flex", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                            <div style={{ fontWeight: 800 }}>{c.name || "—"}</div>
                            <div style={styles.muted}>
                              {c.dateISO || ""} {c.locked ? "🔒" : ""}
                            </div>
                          </div>
                          <div style={{ marginTop: 4, whiteSpace: "pre-wrap" }}>{c.text || ""}</div>
                        </div>
                      ))}
                  </div>
                ) : (
                  <div style={styles.muted}>No comments yet.</div>
                )}
              </div>
            </div>
          </div>
        </div>
      ) : null}
    </div>
  );
}

/* ---------- small components ---------- */
function DatePill({ value, isHeader, overdue, done, muted }) {
  if (isHeader) return <div style={{ color: "#9CA3AF" }}>—</div>;
  const empty = !value;
  return (
    <div
      style={{
        ...styles.pillCompact,
        ...(empty ? styles.pillEmpty : null),
        ...(overdue ? styles.pillLate : null),
        ...(done ? styles.pillDone : null),
        ...(muted ? styles.pillMuted : null),
      }}
    >
      {empty ? "—" : value}
    </div>
  );
}
function MilestoneCell({ isHeader, value, checked, onChange, overdue, disabled, muted, locked, lockMeta, isAdmin, onUnlock }) {
  if (isHeader) return <div style={{ color: "#9CA3AF" }}>—</div>;
  const empty = !value;
  return (
    <div style={styles.milestoneCell}>
      <div
        style={{
          ...styles.pillCompact,
          ...(empty ? styles.pillEmpty : null),
          ...(overdue ? styles.pillLate : null),
          ...(checked ? styles.pillDone : null),
          ...(muted ? styles.pillMuted : null),
          ...(locked ? styles.pillLocked : null),
        }}
      >
        {empty ? "—" : value} {locked ? "🔒" : ""}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <input type="checkbox" checked={!!checked} onChange={(e) => onChange(e.target.checked)} disabled={!!disabled} />
        {locked && isAdmin ? (
          <button style={styles.unlockBtn} onClick={onUnlock} title={lockMeta ? `Locked by ${lockMeta.lockedBy || "—"} on ${lockMeta.lockedAt || ""}` : "Unlock"}>
            Unlock
          </button>
        ) : null}
      </div>
    </div>
  );
}
function StatusBadge({ status }) {
  const map = {
    overdue: { text: "Overdue", style: styles.badgeOverdue },
    ongoing: { text: "Ongoing", style: styles.badgeOngoing },
    done: { text: "Done", style: styles.badgeDone },
  };
  const s = map[status] || { text: status, style: styles.badgeOngoing };
  return <span style={{ ...styles.badgeBase, ...s.style }}>{s.text}</span>;
}
function FilterPill({ label, active, onClick }) {
  return (
    <button onClick={onClick} style={{ ...styles.filterPill, ...(active ? styles.filterPillActive : null) }}>
      {label}
    </button>
  );
}

/* ---------- styles ---------- */
const MAX_W = 1320;
const TEAL = "#0D9488";

const styles = {
  shell: {
    minHeight: "100vh",
    background: "linear-gradient(180deg, #F8FAFC 0%, #F3F4F6 100%)",
  },

  page: {
    fontFamily: "ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial",
    padding: 16,
    margin: "0 auto",
    color: "#111827",
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },

  header: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: 12,
    marginBottom: 12,
    flexWrap: "wrap",
    width: "100%",
    maxWidth: MAX_W,
  },
  headerButtons: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  brandRow: { display: "flex", gap: 12, alignItems: "center" },
  brandLogo: { width: 44, height: 44, objectFit: "contain" },

  savePill: { padding: "8px 12px", borderRadius: 12, border: "1px solid #E5E7EB", fontWeight: 900, fontSize: 12, cursor: "pointer" },
  h1: { fontSize: 22, margin: 0, lineHeight: 1.2, fontWeight: 900 },
  h2: { fontSize: 15, margin: 0, lineHeight: 1.2, fontWeight: 900 },
  h3: { fontSize: 13, margin: 0, lineHeight: 1.2, fontWeight: 900 },
  sub: { margin: "6px 0 0", color: "#4B5563", fontSize: 13 },
  muted: { color: "#6B7280", fontSize: 12 },

    card: {
    background: "rgba(255,255,255,0.9)",
    border: "1px solid #E5E7EB",
    borderRadius: 16,
    padding: 14,
    marginBottom: 12,
    boxShadow: "0 6px 20px rgba(17,24,39,0.06)",
    width: "100%",
    maxWidth: MAX_W,
    backdropFilter: "blur(6px)",
  },
  cardInset: { border: "1px solid #E5E7EB", borderRadius: 14, padding: 10, background: "#FFFFFF" },
  sectionTitle: { fontSize: 12, color: "#6B7280" },
  cardHeader: { display: "flex", alignItems: "baseline", justifyContent: "space-between", gap: 12, marginBottom: 10, flexWrap: "wrap" },

  grid2: { display: "grid", gridTemplateColumns: "repeat(2, minmax(0, 1fr))", gap: 10 },
  label: { display: "grid", gap: 6, fontSize: 13, color: "#374151" },

  input: {
    width: "100%",
    boxSizing: "border-box",
    border: "1px solid #D1D5DB",
    borderRadius: 12,
    padding: "8px 10px",
    fontSize: 12,
    outline: "none",
    background: "#FFFFFF",
  },
  inputMuted: { background: "#F3F4F6", color: "#6B7280" },

  selectorBar: {
    display: "flex",
    gap: 12,
    alignItems: "flex-end",
    justifyContent: "flex-start",
    flexWrap: "wrap",
  },
  selectorBlock: { display: "grid", gap: 6 },

  tableTop: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, marginBottom: 8, flexWrap: "wrap" },
  tableWrap: { overflowX: "auto", WebkitOverflowScrolling: "touch", border: "1px solid #E5E7EB", borderRadius: 14, width: "100%", background: "#fff" },
  table: { width: "100%", minWidth: 980, borderCollapse: "separate", borderSpacing: 0, tableLayout: "fixed" },

  th: { textAlign: "left", fontSize: 12, color: "#374151", background: "#F9FAFB", padding: "10px 10px", borderBottom: "1px solid #E5E7EB" },
  td: { padding: "10px 10px", borderBottom: "1px solid #F3F4F6", verticalAlign: "top", background: "#FFFFFF", fontSize: 12 },

  thSmall: { textAlign: "center", fontSize: 12, color: "#374151", background: "#F9FAFB", padding: "10px 8px", borderBottom: "1px solid #E5E7EB", width: 46 },
  thMed: { textAlign: "left", fontSize: 12, color: "#374151", background: "#F9FAFB", padding: "10px 8px", borderBottom: "1px solid #E5E7EB", width: 115 },
  thWide: { textAlign: "left", fontSize: 12, color: "#374151", background: "#F9FAFB", padding: "10px 8px", borderBottom: "1px solid #E5E7EB", width: 260 },
  thTf: { textAlign: "left", fontSize: 12, color: "#374151", background: "#F9FAFB", padding: "10px 8px", borderBottom: "1px solid #E5E7EB", width: 120 },

  tdCenter: { padding: "10px 8px", borderBottom: "1px solid #F3F4F6", verticalAlign: "top", background: "#FFFFFF", textAlign: "center" },

  trBase: {},

  trHeader: { background: "#F9FAFB" },
  trLate: { background: "#FEF2F2" },
  trDone: { opacity: 0.75 },
  trNotRequired: { background: "#F9FAFB", opacity: 0.65 },

  pillCompact: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, border: "1px solid #E5E7EB", fontSize: 12, background: "#FFFFFF", whiteSpace: "nowrap" },
  pillEmpty: { color: "#9CA3AF", background: "#FAFAFA" },
  pillLate: { border: "1px solid #EF4444" },
  pillDone: { border: "1px solid #10B981" },
  pillMuted: { background: "#F3F4F6", color: "#6B7280" },
  pillLocked: { background: "#F3F4F6", color: "#111827" },

  wrap: { wordBreak: "break-word", whiteSpace: "normal" },

  milestoneCell: { display: "flex", gap: 8, alignItems: "center", justifyContent: "space-between" },

  tfRow: { display: "flex", gap: 6, alignItems: "center" },
  tfInput: { width: 52, border: "1px solid #D1D5DB", borderRadius: 12, padding: "8px 10px", fontSize: 12, outline: "none", background: "#fff" },

  primaryBtn: { background: TEAL, color: "white", border: `1px solid ${TEAL}`, borderRadius: 14, padding: "10px 12px", fontSize: 12, cursor: "pointer", fontWeight: 900 },
  secondaryBtn: { background: "#FFFFFF", color: "#111827", border: "1px solid #D1D5DB", borderRadius: 14, padding: "10px 12px", fontSize: 12, cursor: "pointer", fontWeight: 800 },
  smallBtn: { background: "#FFFFFF", color: "#111827", border: "1px solid #D1D5DB", borderRadius: 12, padding: "8px 10px", fontSize: 12, cursor: "pointer", fontWeight: 800 },
  iconBtn: { background: "#FFFFFF", color: "#111827", border: "1px solid #D1D5DB", borderRadius: 12, padding: "8px 10px", fontSize: 12, cursor: "pointer" },
  unlockBtn: { background: "#FEF2F2", color: "#991B1B", border: "1px solid #FCA5A5", borderRadius: 999, padding: "6px 8px", fontSize: 11, cursor: "pointer", fontWeight: 900 },
  dangerBtn: { background: "#FFCCCB", color: "#111827", border: "1px solid #D1D5DB", borderRadius: 14, padding: "10px 12px", fontSize: 12, cursor: "pointer", fontWeight: 800 },

  /* --- Modal (comments) --- */
  modalOverlay: { position: "fixed", inset: 0, background: "rgba(0,0,0,0.35)", display: "flex", alignItems: "center", justifyContent: "center", padding: 16, zIndex: 50 },
  modalCard: { width: "min(720px, 100%)", background: "#FFFFFF", borderRadius: 18, border: "1px solid #E5E7EB", boxShadow: "0 20px 60px rgba(0,0,0,0.25)", padding: 16 },
  modalHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 },
  commentItem: { border: "1px solid #E5E7EB", borderRadius: 14, padding: 12, background: "#FFFFFF" },

  badgeBase: { display: "inline-flex", alignItems: "center", padding: "6px 10px", borderRadius: 999, fontSize: 12, fontWeight: 900, border: "1px solid transparent", whiteSpace: "nowrap" },
  badgeOverdue: { background: "#FEF2F2", color: "#991B1B", borderColor: "#FCA5A5" },
  badgeOngoing: { background: "#FFFBEB", color: "#92400E", borderColor: "#FDE68A" },
  badgeDone: { background: "#ECFDF5", color: "#065F46", borderColor: "#A7F3D0" },

  filterPill: { background: "#FFFFFF", color: "#111827", border: "1px solid #D1D5DB", borderRadius: 999, padding: "9px 12px", fontSize: 12, cursor: "pointer", fontWeight: 800 },
  filterPillActive: { background: "#111827", color: "#FFFFFF", borderColor: "#111827" },

  stack: { display: "grid", gap: 6 },
  inline: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },

  keyItem: { display: "inline-flex", gap: 8, alignItems: "center" },
  keyText: { fontSize: 12, color: "#374151" },

  /* --- Gantt styles --- */
  ganttWrap: { display: "grid", gap: 8 },
  ganttHeaderRow: { display: "grid", gridTemplateColumns: "260px 1fr 120px 120px", gap: 10, alignItems: "end" },
  ganttRow: { display: "grid", gridTemplateColumns: "260px 1fr 120px 120px", gap: 10, alignItems: "center" },
  ganttHeaderRowDense: { display: "grid", gridTemplateColumns: "260px 1fr 120px 120px", gap: 10, alignItems: "end" },
  ganttRowDense: { display: "grid", gridTemplateColumns: "260px 1fr 120px 120px", gap: 10, alignItems: "center" },
  ganttLabelCol: { minWidth: 0 },
  ganttChartCol: { minWidth: 0 },
  ganttDateCol: { display: "flex", justifyContent: "flex-start" },
  ganttAxis: { position: "relative", height: 34, border: "1px solid #E5E7EB", borderRadius: 12, background: "#F9FAFB", overflow: "hidden" },
  ganttTick: { position: "absolute", top: 0, bottom: 0, transform: "translateX(-0.5px)", pointerEvents: "none" },
  ganttTickLine: { position: "absolute", top: 0, bottom: 0, width: 1, background: "#E5E7EB" },
  ganttTickLabel: { position: "absolute", top: 8, left: 6, fontSize: 11, color: "#6B7280", whiteSpace: "nowrap" },
  ganttLane: { position: "relative", height: 28, border: "1px solid #E5E7EB", borderRadius: 12, background: "#FFFFFF", overflow: "hidden" },
  ganttBar: { position: "absolute", top: 6, height: 16, borderRadius: 999, background: "#111827" },
  ganttBarMissing: { position: "absolute", inset: 0, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 11, color: "#9CA3AF" },

  /* --- Fullscreen programme summary --- */
  fullscreen: { position: "fixed", inset: 0, background: "#F8FAFC", display: "flex", flexDirection: "column" },
  fullTopBar: { padding: 12, background: "#FFFFFF", borderBottom: "1px solid #E5E7EB", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, flexWrap: "wrap" },
  fullBody: { padding: 12, overflow: "auto", display: "flex", justifyContent: "center" },
  fullBodyInner: { width: "100%", maxWidth: MAX_W },
  projectSection: { background: "#FFFFFF", border: "1px solid #E5E7EB", borderRadius: 16, padding: 12, marginBottom: 12, width: "100%", boxShadow: "0 10px 24px rgba(17,24,39,0.06)" },
  projectHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 12, flexWrap: "wrap", marginBottom: 8 },
  projectTitle: { fontSize: 16, fontWeight: 900, marginBottom: 2 },
  projectGantt: { border: "1px solid #E5E7EB", borderRadius: 16, padding: 10, background: "#FFFFFF" },

  /* --- Boot loading --- */
  loadingTopBar: {
    height: 3,
    width: "100%",
    background:
      "linear-gradient(90deg, rgba(16,185,129,0) 0%, rgba(16,185,129,1) 50%, rgba(16,185,129,0) 100%)",
    backgroundSize: "200% 100%",
    animation: "azzureLoading 1.1s linear infinite",
  },
  loadingCard: {
    border: "1px solid #E5E7EB",
    borderRadius: 16,
    background: "#fff",
    padding: 16,
    boxShadow: "0 1px 2px rgba(0,0,0,0.05)",
  },
  loadingTitle: { fontWeight: 900, fontSize: 16, color: "#111827" },
  loadingSub: { marginTop: 6, fontSize: 13, color: "#6B7280" },
};
