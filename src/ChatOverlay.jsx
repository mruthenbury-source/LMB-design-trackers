// src/ChatOverlay.jsx
import React, { useEffect, useMemo, useRef, useState } from "react";

/**
 * ChatOverlay
 * - Always shows a "Chat" launcher button when closed
 * - Full-height right-side panel when open
 * - Header: Search backups (left of Refresh), Refresh, Close (small text)
 * - Footer ALWAYS visible:
 *     1) Quick questions
 *     2) "Ask anything…" input box (free query) BELOW quick questions
 * - Messages area scrolls and never hides footer (minHeight:0 fixes)
 *
 * Supports both prop styles:
 *  A) open/setOpen/messages/input/setInput/sendChat/busy/onReset/searchBackups/setSearchBackups
 *  B) isOpen/onClose/chatMessages/onSend/onRefresh
 */
export default function ChatOverlay(props) {
  const open = props.open ?? props.isOpen ?? false;

  const setOpen =
    props.setOpen ??
    ((v) => {
      if (v === false && typeof props.onClose === "function") props.onClose();
    });

  const messages = props.messages ?? props.chatMessages ?? [];
  const busy = !!(props.busy ?? false);

  // show "Thinking..." while the assistant is generating a response
  const thinking = props.thinking ?? props.isThinking ?? busy;

  const externalInput = props.input;
  const externalSetInput = props.setInput;
  const [localInput, setLocalInput] = useState("");
  const input = externalInput !== undefined ? externalInput : localInput;
  const setInput = externalSetInput !== undefined ? externalSetInput : setLocalInput;

  const externalSearchBackups = props.searchBackups;
  const externalSetSearchBackups = props.setSearchBackups;
  const [localSearchBackups, setLocalSearchBackups] = useState(false);
  const searchBackups =
    externalSearchBackups !== undefined ? externalSearchBackups : localSearchBackups;
  const setSearchBackups =
    externalSetSearchBackups !== undefined ? externalSetSearchBackups : setLocalSearchBackups;

  const onReset = props.onReset ?? props.onRefresh;
  const chatContext = props.chatContext;
  const programmeData = props.programmeData;

const projectOptions = useMemo(() => {
  // Prefer options from props.projects or chatContext.state.projects or chatContext.projects,
  // and fall back to /api/meta results.
  const fromProps = Array.isArray(props.projects) ? props.projects : [];
  const ctx = chatContext || {};
  const fromState = Array.isArray(ctx?.state?.projects) ? ctx.state.projects : [];
  const fromCtxProjects = Array.isArray(ctx?.projects) ? ctx.projects : [];
  const fromMeta = Array.isArray(metaProjects) ? metaProjects : [];

  const names = []
    .concat(fromProps)
    .concat(fromState)
    .concat(fromCtxProjects)
    .concat(fromMeta)
    .map((p) => (typeof p === "string" ? p : p?.name))
    .filter(Boolean)
    .map((s) => String(s).trim())
    .filter(Boolean);

  return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b));
}, [props.projects, chatContext, metaProjects]);const sendChat =
    props.sendChat ??
    (async (text) => {
      if (typeof props.onSend === "function") {
        await props.onSend({ message: text, searchBackups });
      }
    });

  // --- Quick question UI ---
  const [quickId, setQuickId] = useState("");
  const [quickParams, setQuickParams] = useState({ project: "", dateFrom: "", dateTo: "", asOfDate: "" });
const [metaProjects, setMetaProjects] = useState([]);
const [metaLoaded, setMetaLoaded] = useState(false);

useEffect(() => {
  // Pull projects deterministically from /api/meta (optional helper). This keeps the dropdown populated
  // even when chatContext/props don't include the project list.
  if (metaLoaded) return;
  (async () => {
    try {
      const res = await fetch("/api/meta");
      const j = await res.json();
      if (j && j.ok && Array.isArray(j.projects)) setMetaProjects(j.projects);
    } catch {
      // ignore: dropdown will fall back to any projects present in props/chatContext
    } finally {
      setMetaLoaded(true);
    }
  })();
}, [metaLoaded]);
const listRef = useRef(null);

  useEffect(() => {
    if (!open) return;
    const el = listRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [open, messages.length]);

  const programmeHint = useMemo(() => {
    const hasProgramme =
      !!programmeData ||
      (chatContext &&
        (chatContext.programme ||
          chatContext.programmeData ||
          chatContext.programmeIndex ||
          chatContext.programmeRows));
    return hasProgramme
      ? "Use PROGRAMME data (not tracker rows) for on-site level questions. If programme data is missing, say so."
      : "Programme data may be missing—if so, say you cannot answer accurately.";
  }, [chatContext, programmeData]);

  const quickTemplates = useMemo(
  () => [
    { id: "overdue_first_issue_all", label: "Overdue items for first issue (all projects)", params: [] },
    { id: "overdue_first_issue_project", label: "Overdue items for first issue (specific project)", params: [{ key: "project", type: "project" }] },

    { id: "overdue_statusA_all", label: "Overdue items for Status A approval (all projects)", params: [] },
    { id: "overdue_statusA_project", label: "Overdue items for Status A approval (specific project)", params: [{ key: "project", type: "project" }] },

    { id: "statusA_approved_all", label: "Which items have Status A approval (all projects)", params: [] },
    { id: "statusA_approved_project", label: "Which items have Status A approval (specific project)", params: [{ key: "project", type: "project" }] },

    { id: "done_all", label: "Which items are marked as done (all projects)", params: [] },
    { id: "done_project", label: "Which items are marked as done (specific project)", params: [{ key: "project", type: "project" }] },

    {
      id: "programme_range_all",
      label: "Construction programme on site during date range (all projects)",
      params: [{ key: "dateFrom", type: "date" }, { key: "dateTo", type: "date" }],
    },
    {
      id: "programme_range_project",
      label: "Construction programme on site during date range (specific project)",
      params: [{ key: "project", type: "project" }, { key: "dateFrom", type: "date" }, { key: "dateTo", type: "date" }],
    },

    { id: "compare_programme_all", label: "Compare programme: as-of date vs current (all projects, uses backups)", params: [{ key: "asOfDate", type: "date" }] },
    {
      id: "compare_programme_project",
      label: "Compare programme: as-of date vs current (specific project, uses backups)",
      params: [{ key: "project", type: "project" }, { key: "asOfDate", type: "date" }],
    },

    { id: "comments_all", label: "Return all comments (all projects)", params: [] },
    { id: "comments_project", label: "Return comments (specific project)", params: [{ key: "project", type: "project" }] },
  ],
  []
);

    const selectedQuick = useMemo(() => quickTemplates.find((q) => q.id === quickId) || null, [quickTemplates, quickId]);

  function buildQuickPrompt() {
  const q = quickTemplates.find((x) => x.id === quickId);
  if (!q) return "";

  const params = {};
  for (const p of q.params || []) {
    if (p.type === "project") params.project = (quickParams.project || "").trim();
    if (p.type === "date") params[p.key] = (quickParams[p.key] || "").trim();
  }

  for (const p of q.params || []) {
    if (p.type === "project" && !params.project) return "";
    if (p.type === "date" && !params[p.key]) return "";
  }

  // Deterministic marker parsed by the backend:
  // TEMPLATE:<id>\nPARAMS:<json>
  return `TEMPLATE:${q.id}\nPARAMS:${JSON.stringify(params)}`;
}

  function insertQuick() {
    const p = buildQuickPrompt();
    if (!p) return;
    setInput(p);
  }

  async function askQuick() {
    const p = buildQuickPrompt();
    if (!p) return;
    setInput(p);
    await handleSend(p);
  }

  async function handleSend(overrideText) {
    const text = (overrideText ?? input).trim();
    if (!text || busy) return;
    await sendChat(text);
    setInput("");
  }

  // Launcher button when closed so chat never "disappears"
  if (!open) {
    return (
      <button onClick={() => setOpen(true)} style={styles.fab} title="Open chat">
        Chat
      </button>
    );
  }

  return (
    <div style={styles.overlay}>
      <div style={styles.panel}>
        {/* Header */}
        <div style={styles.header}>
          <div style={styles.headerTitleRow}>
          <div style={styles.headerTitle}>Chat</div>
          {thinking ? <div style={styles.thinking}>Thinking…</div> : null}
        </div>

          <div style={styles.headerActions}>
            <label style={styles.topCheckboxRow} title="Include historic weekly backups in search results">
              <input
                type="checkbox"
                checked={!!searchBackups}
                onChange={(e) => setSearchBackups(e.target.checked)}
                style={styles.checkbox}
              />
              <span style={styles.checkboxLabel}>Search backups</span>
            </label>

            <button
              onClick={() => onReset?.()}
              disabled={busy}
              style={{ ...styles.smallBtn, ...(busy ? styles.btnDisabled : null) }}
              title="Refresh chat"
            >
              Refresh
            </button>

            <button onClick={() => setOpen(false)} style={styles.smallBtn} title="Close">
              Close
            </button>
          </div>
        </div>

        {/* Messages (scroll only here) */}
        <div ref={listRef} style={styles.body}>
          {messages.length === 0 ? (
            <div style={styles.empty}>
              Use quick questions below, or ask anything in the query box.
            </div>
          ) : (
            messages.map((m, idx) => {
              const role = m.role ?? m.type;
              const isUser = role === "user";
              const content = m.content ?? m.text ?? "";
              return (
                <div key={idx} style={{ display: "flex", justifyContent: isUser ? "flex-end" : "flex-start" }}>
                  <div style={{ ...styles.bubble, ...(isUser ? styles.userBubble : styles.assistantBubble) }}>
                    {content}
                  </div>
                </div>
              );
            })
          )}
        </div>

        
{/* Footer ALWAYS visible */}
<div style={styles.footer}>
  {/* Quick questions */}
  <div style={styles.quickWrap}>
    <div style={styles.quickRow}>
      <select
        value={quickId}
        onChange={(e) => {
          setQuickId(e.target.value);
          setQuickParams({ project: "", dateFrom: "", dateTo: "", asOfDate: "" });
        }}
        style={styles.select}
      >
        <option value="">Quick questions…</option>
        {quickTemplates.map((q) => (
          <option key={q.id} value={q.id}>
            {q.label}
          </option>
        ))}
      </select>

      <button
        type="button"
        onClick={insertQuick}
        disabled={!buildQuickPrompt()}
        style={{ ...styles.smallBtn, ...(!buildQuickPrompt() ? styles.btnDisabled : null) }}
      >
        Insert
      </button>

      <button
        type="button"
        onClick={askQuick}
        disabled={!buildQuickPrompt()}
        style={{ ...styles.smallBtn, ...(!buildQuickPrompt() ? styles.btnDisabled : null) }}
      >
        Ask
      </button>
    </div>

    {selectedQuick?.params?.length > 0 && (
      <div style={styles.quickRow}>
        {selectedQuick.params.some((p) => p.type === "project") && (
          <select
            value={quickParams.project}
            onChange={(e) => setQuickParams((s) => ({ ...s, project: e.target.value }))}
            style={styles.select}
          >
            <option value="">{projectOptions.length ? "Select project…" : (metaLoaded ? "No projects found" : "Loading projects…")}</option>
            {projectOptions.map((p) => (
              <option key={p} value={p}>
                {p}
              </option>
            ))}
          </select>
        )}

        {selectedQuick.params.some((p) => p.key === "dateFrom") && (
          <input
            type="date"
            value={quickParams.dateFrom}
            onChange={(e) => setQuickParams((s) => ({ ...s, dateFrom: e.target.value }))}
            style={styles.input}
            title="From"
          />
        )}

        {selectedQuick.params.some((p) => p.key === "dateTo") && (
          <input
            type="date"
            value={quickParams.dateTo}
            onChange={(e) => setQuickParams((s) => ({ ...s, dateTo: e.target.value }))}
            style={styles.input}
            title="To"
          />
        )}

        {selectedQuick.params.some((p) => p.key === "asOfDate") && (
          <input
            type="date"
            value={quickParams.asOfDate}
            onChange={(e) => setQuickParams((s) => ({ ...s, asOfDate: e.target.value }))}
            style={styles.input}
            title="As-of date"
          />
        )}
      </div>
    )}
  </div>

  {/* Query box */}
  <div style={styles.queryBlock}>
    <div style={styles.queryBoxLabel}>Query</div>

    <div style={styles.composer}>
      <textarea
        value={input}
        onChange={(e) => setInput(e.target.value)}
        placeholder="Type a question, or use Quick questions above…"
        style={styles.textarea}
        disabled={busy}
      />
    </div>

    <div style={styles.sendRow}>
      <button
        type="button"
        onClick={() => handleSend()}
        disabled={busy || !input.trim()}
        style={{ ...styles.sendBtn, ...(busy || !input.trim() ? styles.btnDisabled : null) }}
      >
        Send
      </button>
    </div>
  </div>
</div>
      </div>
    </div>
  );
}

const styles = {
  fab: {
    position: "fixed",
    right: 16,
    bottom: 16,
    zIndex: 9999,
    borderRadius: 999,
    border: "1px solid #111827",
    background: "#111827",
    color: "#fff",
    padding: "10px 14px",
    fontSize: 13,
    fontWeight: 700,
    cursor: "pointer",
    boxShadow: "0 10px 24px rgba(17,24,39,0.25)",
  },

  overlay: {
    position: "fixed",
    top: 0,
    right: 0,
    height: "100dvh",
    // Uses dynamic viewport height on mobile

    width: 420,
    maxWidth: "95vw",
    zIndex: 9999,
    pointerEvents: "auto",
  },
  panel: {
    height: "100%",
    minHeight: 0,
    overflow: "hidden",
    boxSizing: "border-box", // IMPORTANT so body can scroll and footer stays visible
    background: "#fff",
    borderLeft: "1px solid #e5e7eb",
    boxShadow: "0 16px 40px rgba(0,0,0,0.16)",
    display: "flex",
    flexDirection: "column",
  },

  header: {
    flexShrink: 0,
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 10,
    padding: "10px 12px",
    borderBottom: "1px solid #e5e7eb",
    background: "#fff",
  },
  headerTitle: { fontWeight: 700, fontSize: 13, color: "#111827" },
  headerActions: { display: "flex", gap: 8, alignItems: "center" },

  topCheckboxRow: {
    display: "flex",
    gap: 6,
    alignItems: "center",
    border: "1px solid #e5e7eb",
    borderRadius: 8,
    padding: "4px 8px",
    background: "#fff",
  },
  checkbox: { width: 12, height: 12 },
  checkboxLabel: { fontSize: 11, color: "#111827" },

  smallBtn: {
    flexShrink: 0,
    fontSize: 11,
    padding: "4px 8px",
    borderRadius: 8,
    border: "1px solid #e5e7eb",
    background: "#fff",
    cursor: "pointer",
    lineHeight: 1.2,
  },
  btnDisabled: { opacity: 0.5, cursor: "not-allowed" },

  body: {
    flex: 1,
    minHeight: 0, // IMPORTANT for scrolling in flex layouts
    overflowY: "auto",
    padding: 12,
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  empty: { fontSize: 13, color: "#6b7280" },
  bubble: {
    maxWidth: "85%",
    whiteSpace: "pre-wrap",
    borderRadius: 12,
    padding: "8px 10px",
    fontSize: 13,
    lineHeight: 1.35,
  },
  userBubble: { background: "#111827", color: "#fff" },
  assistantBubble: { background: "#f3f4f6", color: "#111827" },

  footer: {
    flexShrink: 0,
    position: "sticky",
    bottom: 0,
    left: 0,
    right: 0,
    borderTop: "1px solid #e5e7eb",
    padding: 10,
    paddingBottom: "calc(18px + env(safe-area-inset-bottom))",
    display: "flex",
    flexDirection: "column",
    gap: 10,
    background: "#fff",
  },

  quickWrap: { display: "flex", flexDirection: "column", gap: 8 },
  quickRow: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  select: {
    flex: "1 1 260px",
    minWidth: 220,
    border: "1px solid #e5e7eb",
    borderRadius: 8,
    padding: "6px 8px",
    fontSize: 13,
    background: "#fff",
  },
  input: {
    flex: 1,
    border: "1px solid #e5e7eb",
    borderRadius: 8,
    padding: "6px 8px",
    fontSize: 13,
  },

  queryBoxLabel: {
    fontSize: 12,
    fontWeight: 700,
    color: "#111827",
    marginTop: 2,
  },
queryBlock: {
  display: "flex",
  flexDirection: "column",
  gap: 8,
},
sendRow: {
  display: "flex",
  justifyContent: "flex-end",
},

  composer: {
    display: "flex",
    gap: 8,
    alignItems: "stretch",
    flexShrink: 0,
  },
  textarea: {
    width: "100%",
    resize: "none",
    background: "#ffffff",
    border: "2px solid #e5e7eb",
    borderRadius: 10,
    padding: 8,
    fontSize: 13,
    lineHeight: 1.3,
    minHeight: 70,
    maxHeight: 160,
    overflowY: "auto",
    boxSizing: "border-box",
  },
  sendBtn: {
    width: "100%",
    maxWidth: 140,
    alignSelf: "flex-end",
    borderRadius: 10,
    border: "1px solid #111827",
    background: "#111827",
    color: "#fff",
    fontSize: 13,
    padding: "10px 12px",
    cursor: "pointer",
    boxSizing: "border-box",
  },
};
