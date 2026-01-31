// src/ChatOverlay.jsx
import React, { useEffect, useMemo, useRef, useState } from "react";

/**
 * ChatOverlay (full file)
 *
 * Goals implemented:
 * - Full-height right-side overlay (visible on every page where mounted)
 * - Header controls small text: "Search backups" (LEFT of Refresh), Refresh, Close
 * - Quick questions block in footer
 * - Free-typing chat input ALWAYS visible BELOW quick questions
 * - Message list scrolls (does not push footer off screen)
 * - "All levels on site…" quick questions explicitly reference PROGRAMME data (not tracker)
 *
 * Compatibility:
 * This component supports BOTH prop styles used in your codebase:
 *
 * A) "open" style:
 *    <ChatOverlay open={open} setOpen={setOpen} messages={messages} input={input} setInput={setInput}
 *                sendChat={sendChat} busy={busy} onReset={onReset}
 *                searchBackups={searchBackups} setSearchBackups={setSearchBackups}
 *                chatContext={chatContext} programmeData={programmeData} />
 *
 * B) "isOpen" style:
 *    <ChatOverlay isOpen={isOpen} onClose={onClose} chatMessages={chatMessages} onSend={onSend}
 *                onRefresh={onRefresh} />
 *
 * If your app passes only one style, this will still work.
 */

export default function ChatOverlay(props) {
  // ---- Normalise props across both call sites ----
  const open = props.open ?? props.isOpen ?? false;

  const setOpen =
    props.setOpen ??
    ((v) => {
      // fallback to onClose if present
      if (v === false && typeof props.onClose === "function") props.onClose();
    });

  const messages = props.messages ?? props.chatMessages ?? [];
  const busy = props.busy ?? false;

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

  // send function: prefer sendChat(text) OR onSend({message, searchBackups})
  const sendChat =
    props.sendChat ??
    (async (text) => {
      if (typeof props.onSend === "function") {
        await props.onSend({ message: text, searchBackups });
      }
    });

  // ---- Quick question UI state ----
  const [quickId, setQuickId] = useState("");
  const [quickProject, setQuickProject] = useState("");
  const [quickOnSite, setQuickOnSite] = useState("");

  const listRef = useRef(null);

  // Auto-scroll to bottom on new messages
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
            "If programme dates are missing/ambiguous, state what’s missing and which fields you checked.",
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
            "If the project name doesn't match exactly, show the closest matches and ask which to use.",
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
    const prompt = buildQuickPrompt();
    if (!prompt) return;
    setInput(prompt);
  }

  async function askQuick() {
    const prompt = buildQuickPrompt();
    if (!prompt) return;
    setInput(prompt);
    await handleSend(prompt);
  }

  async function handleSend(overrideText) {
    const text = (overrideText ?? input).trim();
    if (!text || busy) return;
    await sendChat(text);
    setInput("");
  }

  if (!open) return null;

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

        {/* Body (scrollable messages) */}
        <div ref={listRef} style={styles.body}>
          {messages.length === 0 ? (
            <div style={styles.empty}>
              Ask anything below, or use a quick question.
            </div>
          ) : (
            messages.map((m, idx) => {
              const role = m.role ?? m.type; // support both shapes
              const isUser = role === "user";
              const content = m.content ?? m.text ?? "";
              return (
                <div
                  key={idx}
                  style={{
                    display: "flex",
                    justifyContent: isUser ? "flex-end" : "flex-start",
                  }}
                >
                  <div style={{ ...styles.bubble, ...(isUser ? styles.userBubble : styles.assistantBubble) }}>
                    {content}
                  </div>
                </div>
              );
            })
          )}
        </div>

        {/* Footer: Quick questions (top) + chat input (below) */}
        <div style={styles.footer}>
          {/* Quick questions */}
          <div style={styles.quickWrap}>
            <div style={styles.quickRow}>
              <select
                value={quickId}
                onChange={(e) => setQuickId(e.target.value)}
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
                style={{ ...styles.smallBtn, ...(buildQuickPrompt() ? null : styles.btnDisabled) }}
                title="Insert into message box"
              >
                Insert
              </button>

              <button
                type="button"
                onClick={askQuick}
                disabled={!buildQuickPrompt()}
                style={{ ...styles.smallBtn, ...(buildQuickPrompt() ? null : styles.btnDisabled) }}
                title="Send immediately"
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

          {/* Chat input (ALWAYS visible) */}
          <div style={styles.composer}>
            <textarea
              value={input}
              onChange={(e) => setInput(e.target.value)}
              placeholder="Ask any question…"
              rows={3}
              style={styles.textarea}
              onKeyDown={(e) => {
                // Enter to send (Shift+Enter for newline)
                if (e.key === "Enter" && !e.shiftKey) {
                  e.preventDefault();
                  handleSend();
                }
              }}
            />
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
  );
}

const styles = {
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
    background: "#fff",
    borderLeft: "1px solid #e5e7eb",
    boxShadow: "0 16px 40px rgba(0,0,0,0.16)",
    display: "flex",
    flexDirection: "column",
  },
  header: {
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
    borderTop: "1px solid #e5e7eb",
    padding: 10,
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

  composer: {
    display: "flex",
    gap: 8,
    alignItems: "stretch",
  },
  textarea: {
    flex: 1,
    resize: "none",
    border: "1px solid #e5e7eb",
    borderRadius: 10,
    padding: 8,
    fontSize: 13,
    lineHeight: 1.3,
    minHeight: 70,
  },
  sendBtn: {
    width: 72,
    borderRadius: 10,
    border: "1px solid #111827",
    background: "#111827",
    color: "#fff",
    fontSize: 13,
    cursor: "pointer",
  },
};
