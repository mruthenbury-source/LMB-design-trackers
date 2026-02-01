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

  // --- Quick question UI ---
  const [quickId, setQuickId] = useState("");
  const [quickProject, setQuickProject] = useState("");
  const [quickOnSite, setQuickOnSite] = useState("");

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
      {
        id: "statusA_all",
        label: "Status A overdue — all projects",
        needsProject: false,
        needsOnSite: false,
        build: () =>
          [
            "Using ONLY the current tracker context, list all projects with Status A overdue.",
            "Return: Project name, Level/Block, item/row reference, due date, days overdue.",
            "If a field is missing, state it—do not guess.",
          ].join("\n"),
      },
      {
        id: "statusA_project",
        label: "Status A overdue — specific project",
        needsProject: true,
        needsOnSite: false,
        build: (p) =>
          [
            `Using ONLY the current tracker context, list all Status A overdue items for project: "${p}".`,
            "Return: Level/Block, item/row reference, due date, days overdue.",
            "If the project name doesn't match exactly, show the closest matches and ask which to use.",
          ].join("\n"),
      },
      {
        id: "firstIssue_all",
        label: "First Issue overdue — all projects",
        needsProject: false,
        needsOnSite: false,
        build: () =>
          [
            "Using ONLY the current tracker context, list all projects with First Issue overdue.",
            "Return: Project name, Level/Block, item/row reference, due date, days overdue.",
            "If a field is missing, state it—do not guess.",
          ].join("\n"),
      },
      {
        id: "firstIssue_project",
        label: "First Issue overdue — specific project",
        needsProject: true,
        needsOnSite: false,
        build: (p) =>
          [
            `Using ONLY the current tracker context, list all First Issue overdue items for project: "${p}".`,
            "Return: Level/Block, item/row reference, due date, days overdue.",
            "If the project name doesn't match exactly, show the closest matches and ask which to use.",
          ].join("\n"),
      },
      {
        id: "onsite_all",
        label: "All levels on site in… — all projects (programme)",
        needsProject: false,
        needsOnSite: true,
        build: (_p, d) =>
          [
            programmeHint,
            `Using ONLY the PROGRAMME context, list ALL levels across ALL projects that are on site in: "${d}".`,
            "Interpret the input as a month/date/range and match levels scheduled on site within that period.",
            "Return: Project name, Level/Block, on-site date (or range), and which programme field(s) you used.",
            "Do NOT use tracker rows for this question.",
          ].join("\n"),
      },
      {
        id: "onsite_project",
        label: "All levels on site in… — specific project (programme)",
        needsProject: true,
        needsOnSite: true,
        build: (p, d) =>
          [
            programmeHint,
            `Using ONLY the PROGRAMME context, list ALL levels for project: "${p}" that are on site in: "${d}".`,
            "Interpret the input as a month/date/range and match levels scheduled on site within that period.",
            "Return: Level/Block, on-site date (or range), and which programme field(s) you used.",
            "Do NOT use tracker rows for this question.",
          ].join("\n"),
      },
    ],
    [programmeHint]
  );

  const selectedQuick = useMemo(
    () => quickTemplates.find((q) => q.id === quickId) || null,
    [quickId, quickTemplates]
  );

  function buildQuickPrompt() {
    if (!selectedQuick) return "";
    if (selectedQuick.needsProject && !quickProject.trim()) return "";
    if (selectedQuick.needsOnSite && !quickOnSite.trim()) return "";
    return selectedQuick.build(quickProject.trim(), quickOnSite.trim());
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
          <div style={styles.headerTitle}>Chat</div>

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
              <select value={quickId} onChange={(e) => setQuickId(e.target.value)} style={styles.select}>
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

            {(selectedQuick?.needsProject || selectedQuick?.needsOnSite) && (
              <div style={styles.quickRow}>
                {selectedQuick?.needsProject && (
                  <input
                    value={quickProject}
                    onChange={(e) => setQuickProject(e.target.value)}
                    placeholder="Project name…"
                    style={styles.input}
                  />
                )}
                {selectedQuick?.needsOnSite && (
                  <input
                    value={quickOnSite}
                    onChange={(e) => setQuickOnSite(e.target.value)}
                    placeholder="Month / date / range…"
                    style={styles.input}
                  />
                )}
              </div>
            )}
          </div>

{/* FREE QUERY BOX (always visible, below quick questions) */}
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
    height: "100vh",
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
    borderTop: "1px solid #e5e7eb",
    padding: 10,
    paddingBottom: "calc(10px + env(safe-area-inset-bottom))",
    display: "flex",
    flexDirection: "column",
    gap: 10,
    background: "#fff",
  },

  quickWrap: { display: "flex", flexDirection: "column", gap: 8 },
  quickRow: { display: "flex", gap: 8, alignItems: "center" },
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
