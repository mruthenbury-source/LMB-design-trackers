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

  const sendChat =
    props.sendChat ??
    (async (text) => {
      if (typeof props.onSend === "function") {
        await props.onSend({ message: text, searchBackups });
      }
    });

  const [lastChatResult, setLastChatResult] = useState(null);

function getLastAssistantText() {
  const arr = Array.isArray(messages) ? messages : [];
  for (let i = arr.length - 1; i >= 0; i--) {
    const m = arr[i];
    if (m?.role === "assistant") {
      if (typeof m.content === "string") return m.content;
      if (typeof m.text === "string") return m.text;
    }
  }
  return "";
}

async function exportPdf() {
  const answerText = (lastChatResult?.answer || getLastAssistantText() || "").trim();
  if (!answerText) return;

  const payload = {
    title: "LMD SupplySync Report",
    answer: answerText,
    rows: Array.isArray(lastChatResult?.data?.rows) ? lastChatResult.data.rows : [],
    meta: lastChatResult?.data?.meta || {},
  };

  const res = await fetch("/api/exportPdf", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    const t = await res.text().catch(() => "");
    alert("Export failed. " + (t || res.status));
    return;
  }

  const blob = await res.blob();
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "chat-export.pdf";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}


  // --- Quick question UI ---
  const [quickId, setQuickId] = useState("");
const [quickProject, setQuickProject] = useState("");
const [quickOnSite, setQuickOnSite] = useState(""); // kept for backward compatibility (unused by templates)
const [quickDateFrom, setQuickDateFrom] = useState("");
const [quickDateTo, setQuickDateTo] = useState("");
const [quickAsOfDate, setQuickAsOfDate] = useState("");

const [metaProjects, setMetaProjects] = useState([]);
const [metaLoaded, setMetaLoaded] = useState(false);

useEffect(() => {
  if (metaLoaded) return;
  (async () => {
    try {
      const res = await fetch("/api/meta");
      const j = await res.json();
      if (j && j.ok && Array.isArray(j.projects)) setMetaProjects(j.projects);
    } catch {
      // ignore
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
    { id: "overdue_first_issue_all", label: "Overdue items for first issue on all projects", needsProject: false, needsDateRange: false, needsAsOf: false },
    { id: "overdue_first_issue_project", label: "Overdue items for first issue on a specific project", needsProject: true, needsDateRange: false, needsAsOf: false },

    { id: "overdue_statusA_all", label: "Overdue items for Status A approval (all projects)", needsProject: false, needsDateRange: false, needsAsOf: false },
    { id: "overdue_statusA_project", label: "Overdue items for Status A approval on a specific project", needsProject: true, needsDateRange: false, needsAsOf: false },

    { id: "statusA_approved_all", label: "Which items have Status A approval for all projects", needsProject: false, needsDateRange: false, needsAsOf: false },
    { id: "statusA_approved_project", label: "Which items have Status A approval for a specific project", needsProject: true, needsDateRange: false, needsAsOf: false },

    { id: "done_all", label: "Which items are marked as done for all projects", needsProject: false, needsDateRange: false, needsAsOf: false },
    { id: "done_project", label: "Which items are marked as done for a specific project", needsProject: true, needsDateRange: false, needsAsOf: false },

    { id: "programme_range_all", label: "Construction programme for ALL projects on site during (date range)", needsProject: false, needsDateRange: true, needsAsOf: false },
    { id: "programme_range_project", label: "Construction programme for a specific project on site during (date range)", needsProject: true, needsDateRange: true, needsAsOf: false },

    { id: "compare_programme_all", label: "Compare construction programmes for ALL projects (as-of date vs current)", needsProject: false, needsDateRange: false, needsAsOf: true },
    { id: "compare_programme_project", label: "Compare construction programme for a specific project (as-of date vs current)", needsProject: true, needsDateRange: false, needsAsOf: true },

    { id: "comments_all", label: "Return all comments for all projects", needsProject: false, needsDateRange: false, needsAsOf: false },
    { id: "comments_project", label: "Return comments for a specific project", needsProject: true, needsDateRange: false, needsAsOf: false },
  ],
  []
);

const projectOptions = useMemo(() => {
  const ctx = chatContext || {};
  const fromCtxState = Array.isArray(ctx?.state?.projects) ? ctx.state.projects : [];
  const fromCtxProjects = Array.isArray(ctx?.projects) ? ctx.projects : [];
  const fromProps = Array.isArray(props.projects) ? props.projects : [];
  const names = []
    .concat(fromProps)
    .concat(fromCtxState)
    .concat(fromCtxProjects)
    .concat(Array.isArray(metaProjects) ? metaProjects : [])
    .map((p) => (typeof p === "string" ? p : p?.name))
    .filter(Boolean)
    .map((s) => String(s).trim())
    .filter(Boolean);
  return Array.from(new Set(names)).sort((a, b) => a.localeCompare(b));
}, [props.projects, chatContext, metaProjects]);
  const selectedQuick = useMemo(
    () => quickTemplates.find((q) => q.id === quickId) || null,
    [quickId, quickTemplates]
  );

function buildQuickPrompt() {
  const q = quickTemplates.find((x) => x.id === quickId);
  if (!q) return "";

  const params = {};

  if (q.needsProject) {
    const p = String(quickProject || "").trim();
    if (!p) return "";
    params.project = p;
  }

  if (q.needsDateRange) {
    const from = String(quickDateFrom || "").trim();
    const to = String(quickDateTo || "").trim();
    if (!from || !to) return "";
    params.dateFrom = from;
    params.dateTo = to;
  }

  if (q.needsAsOf) {
    const asOf = String(quickAsOfDate || "").trim();
    if (!asOf) return "";
    params.asOfDate = asOf;
  }

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
            (Array.isArray(messages) ? messages : [])
  .filter(Boolean)
  .map((m, idx) => {
    const role = m?.role ?? m?.type;
    const isUser = role === "user";
    const raw = m?.content ?? m?.text ?? "";
    const content =
      typeof raw === "string" ? raw : (() => { try { return JSON.stringify(raw); } catch { return String(raw); } })();

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
        setQuickProject("");
        setQuickDateFrom("");
        setQuickDateTo("");
        setQuickAsOfDate("");
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

  {(selectedQuick?.needsProject || selectedQuick?.needsDateRange || selectedQuick?.needsAsOf) && (
    <div style={styles.quickRow}>
      {selectedQuick?.needsProject && (
        <select
          value={quickProject}
          onChange={(e) => setQuickProject(e.target.value)}
          style={styles.select}
        >
          <option value="">
            {projectOptions.length ? "Select project…" : (metaLoaded ? "No projects found" : "Loading projects…")}
          </option>
          {projectOptions.map((p) => (
            <option key={p} value={p}>
              {p}
            </option>
          ))}
        </select>
      )}

      {selectedQuick?.needsDateRange && (
        <>
          <input
            type="date"
            value={quickDateFrom}
            onChange={(e) => setQuickDateFrom(e.target.value)}
            style={styles.input}
            title="From"
          />
          <input
            type="date"
            value={quickDateTo}
            onChange={(e) => setQuickDateTo(e.target.value)}
            style={styles.input}
            title="To"
          />
        </>
      )}

      {selectedQuick?.needsAsOf && (
        <input
          type="date"
          value={quickAsOfDate}
          onChange={(e) => setQuickAsOfDate(e.target.value)}
          style={styles.input}
          title="As-of date"
        />
      )}
    </div>
  )}
</div>

          {/* FREE QUERY BOX (always visible) */}
          <div style={styles.queryBlock}>
            <div style={styles.queryBoxLabel}>Ask anything:</div>

  <textarea
    value={input}
    onChange={(e) => setInput(e.target.value)}
    placeholder="Type a message…"
    style={styles.textarea}
    rows={2}
    disabled={false}
    onKeyDown={(e) => {
      // Enter to send, Shift+Enter newline
      if (e.key === "Enter" && !e.shiftKey) {
        e.preventDefault();
        handleSend();
      }
    }}
  />

  <div style={styles.sendRow}>
<button
  type="button"
  onClick={exportPdf}
  disabled={!((lastChatResult?.answer || getLastAssistantText() || "").trim())}
  style={{
    ...styles.smallBtn,
    marginRight: 8,
    ...(!((lastChatResult?.answer || getLastAssistantText() || "").trim()) ? styles.btnDisabled : null),
  }}
  title="Export the latest answer to PDF"
>
  Export PDF
</button>
<button
      onClick={() => handleSend()}
      disabled={busy || !String(input).trim()}
      style={{ ...styles.sendBtn, ...(busy || !String(input).trim() ? styles.btnDisabled : null) }}
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

  smallBtn: { flexShrink: 0,
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
    flex: 1,
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
