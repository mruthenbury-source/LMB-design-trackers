import React, { useEffect, useMemo, useState } from "react";

export default function ChatOverlay({
  open,
  setOpen,
  messages,
  input,
  setInput,
  searchBackups,
  setSearchBackups,
  busy,
  sendChat,
  endRef,
  status, // "idle" | "checking" | "ok" | "error"
  onReset, // ✅ NEW: function passed in from App
}) {
  const [quickId, setQuickId] = useState("statusA_all");
  const [quickProject, setQuickProject] = useState("");
  const [quickOnSite, setQuickOnSite] = useState("");

  const quickTemplates = useMemo(
    () =>
      ({
        statusA_all: {
          label: "Status A overdue — all projects",
          needsProject: false,
          needsOnSite: false,
          build: () =>
            "From the current tracker context ONLY, list every row that is **Status A overdue** across ALL projects. Treat a row as overdue when Required On Site is in the past and Status A is missing/blank OR Status A occurs after Required On Site. Exclude rows marked done. Output a table with: Project, Level/Block, Item, Required On Site, Status A, Days overdue, Owner/Supplier (if present).",
        },
        statusA_project: {
          label: "Status A overdue — specific project",
          needsProject: true,
          needsOnSite: false,
          build: (p) =>
            `From the current tracker context ONLY, list every row that is **Status A overdue** for the project named "${p}". Use the same overdue logic as above and exclude rows marked done. Output a table with: Level/Block, Item, Required On Site, Status A, Days overdue, Owner/Supplier (if present). If the project name is slightly different in the data, use the closest match and tell me what you matched.`,
        },
        firstIssue_all: {
          label: "First Issue overdue — all projects",
          needsProject: false,
          needsOnSite: false,
          build: () =>
            "From the current tracker context ONLY, list every row that is **First Issue overdue** across ALL projects. Treat a row as overdue when the First Issue date is in the past and the row is not done. Output a table with: Project, Level/Block, Item, First Issue, Days overdue, Owner/Supplier (if present).",
        },
        firstIssue_project: {
          label: "First Issue overdue — specific project",
          needsProject: true,
          needsOnSite: false,
          build: (p) =>
            `From the current tracker context ONLY, list every row that is **First Issue overdue** for the project named "${p}". Treat a row as overdue when the First Issue date is in the past and the row is not done. Output a table with: Level/Block, Item, First Issue, Days overdue, Owner/Supplier (if present). If the project name is slightly different in the data, use the closest match and tell me what you matched.`,
        },
        onsite_all: {
          label: "All levels on site in… — all projects",
          needsProject: false,
          needsOnSite: true,
          build: (_p, d) =>
            `From the current programme context ONLY (use programme/programmeIndex, NOT tracker rows), list ALL levels/blocks across ALL projects that are **on site in**: "${d}". (The input may be a month like "March 2026" or a date range; interpret it carefully and tell me what window you used.) Output a table with: Project, Block/Zone, Level, Start Date, Finish Date, On-site range.`,
        },
        onsite_project: {
          label: "All levels on site in… — specific project",
          needsProject: true,
          needsOnSite: true,
          build: (p, d) =>
            `From the current programme context ONLY (use programme/programmeIndex, NOT tracker rows), list ALL levels/blocks for the project named "${p}" that are **on site in**: "${d}". Interpret the date/month input carefully and tell me what window you used. Output a table with: Block/Zone, Level, Start Date, Finish Date, On-site range. If the project name is slightly different in the data, use the closest match and tell me what you matched.`,
        },
      }),
    []
  );

  function applyQuick({ autoSend } = { autoSend: false }) {
    const t = quickTemplates[quickId];
    if (!t) return;
    const project = (quickProject || "").trim();
    const onSite = (quickOnSite || "").trim();

    // Build prompt with required inputs (but don't block the UI—fallback to placeholders)
    const p = t.needsProject ? (project || "<enter project name>") : "";
    const d = t.needsOnSite ? (onSite || "<enter month/date/range>") : "";

    const prompt = t.build(p, d);
    setInput(prompt);

    // Optionally send immediately
    if (autoSend && sendChat && prompt.trim()) {
      // Give React a tick to apply setInput before send
      setTimeout(() => sendChat(), 0);
    }
  }
  // close on ESC
  useEffect(() => {
    if (!open) return;
    const onKey = (e) => {
      if (e.key === "Escape") setOpen(false);
      // Cmd/Ctrl + Enter to send
      if ((e.metaKey || e.ctrlKey) && e.key === "Enter") sendChat?.();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [open, setOpen, sendChat]);

  if (!open) {
    return (
      <button
        onClick={() => setOpen(true)}
        style={styles.fab}
        title="Open chat"
      >
        Chat
      </button>
    );
  }

  const showError = status === "error";

  return (
    <div style={styles.overlay}>
      <div style={styles.panel}>
        {/* Top bar */}
        <div style={styles.topbar}>
          <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
            <div style={styles.title}>Assistant</div>
            <div style={styles.subtleStatus}>
              {busy ? "Thinking…" : status === "ok" ? "Ready" : " "}
            </div>
          </div>

          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            <label
              style={styles.topCheckboxRow}
              title="Include historic weekly backups in search results"
            >
              <input
                type="checkbox"
                checked={!!searchBackups}
                onChange={(e) => setSearchBackups?.(e.target.checked)}
                style={styles.checkbox}
              />
              <span style={styles.checkboxLabel}>Search backups</span>
            </label>

            <button
              onClick={() => onReset?.()}
              disabled={busy}
              style={{
                ...styles.smallBtn,
                ...(busy ? styles.btnDisabled : null),
              }}
              title="Clear chat and start fresh"
            >
              Refresh
            </button>

            <button
              onClick={() => setOpen(false)}
              style={styles.smallBtn}
              title="Close"
            >
              Close
            </button>
          </div>
        </div>

        {/* Body */}
        <div style={styles.body}>
          {messages.map((m, idx) => {
            const isUser = m.role === "user";
            return (
              <div
                key={idx}
                style={{
                  display: "flex",
                  justifyContent: isUser ? "flex-end" : "flex-start",
                  marginBottom: 8,
                }}
              >
                <div
                  style={{
                    ...styles.bubble,
                    ...(isUser ? styles.userBubble : styles.assistantBubble),
                  }}
                >
                  {String(m.content || "")}
                </div>
              </div>
            );
          })}

          {/* Subtle error line (only if truly error) */}
          {showError ? (
            <div style={styles.errorLine}>
              Couldn’t reach the chat server. You can keep working, or press
              Refresh to reset the chat.
            </div>
          ) : null}

          <div ref={endRef} />
        </div>

        {/* Input */}
        <div style={styles.footer}>

  {/* Quick Questions */}
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

      <button onClick={insertQuick} style={styles.smallBtn}>
        Insert
      </button>

      <button onClick={askQuick} style={styles.smallBtn}>
        Ask
      </button>
    </div>
  </div>

  {/* ALWAYS BELOW QUICK QUESTIONS */}
  <div style={styles.queryBlock}>
    <div style={styles.queryLabel}>Ask anything:</div>

    <textarea
      value={input}
      onChange={(e) => setInput(e.target.value)}
      placeholder="Type a message…"
      style={styles.textarea}
      rows={2}
      disabled={false}
      onKeyDown={(e) => {
        if (e.key === "Enter" && !e.shiftKey) {
          e.preventDefault();
          sendChat?.();
        }
      }}
    />

    <button onClick={() => sendChat?.()} style={styles.sendBtn}>
      Send
    </button>
  </div>

</div>


const styles = {
  fab: {
    position: "fixed",
    right: 18,
    bottom: 18,
    zIndex: 50,
    borderRadius: 999,
    padding: "12px 14px",
    border: "1px solid #D1D5DB",
    background: "#111827",
    color: "white",
    fontWeight: 900,
    cursor: "pointer",
    boxShadow: "0 12px 30px rgba(17,24,39,0.25)",
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
  background: "#fff",
  borderLeft: "1px solid #e5e7eb",
  boxShadow: "0 16px 40px rgba(0,0,0,0.16)",
  display: "flex",
  flexDirection: "column",
},


  quickWrap: {
    padding: "10px 10px 0 10px",
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  quickRow: {
    display: "flex",
    gap: 8,
    alignItems: "center",
  },
  quickSelect: {
    flex: 1,
    border: "1px solid #E5E7EB",
    borderRadius: 10,
    padding: "8px 10px",
    background: "white",
    fontSize: 12,
  },
  quickBtn: {
    border: "1px solid #E5E7EB",
    borderRadius: 10,
    padding: "8px 10px",
    background: "#F9FAFB",
    cursor: "pointer",
    fontWeight: 800,
    fontSize: 12,
    whiteSpace: "nowrap",
  },
  quickInput: {
    width: "100%",
    border: "1px solid #E5E7EB",
    borderRadius: 10,
    padding: "8px 10px",
    background: "white",
    fontSize: 12,
  },

  topCheckboxRow: {
  display: "flex",
  gap: 6,
  alignItems: "center",
  border: "1px solid #e5e7eb",
  borderRadius: 8,
  padding: "4px 8px",
  background: "#fff",
},


  topbar: {
    padding: 12,
    borderBottom: "1px solid #E5E7EB",
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    background: "rgba(249,250,251,0.9)",
  },

  title: { fontWeight: 900, fontSize: 13, color: "#111827" },
  subtleStatus: { fontSize: 12, color: "#6B7280" },

  body: {
padding: 12,
    overflow: "auto",
    flex: 1,
  },


  bubble: {
    maxWidth: "85%",
    padding: "10px 12px",
    borderRadius: 14,
    fontSize: 12,
    whiteSpace: "pre-wrap",
    lineHeight: 1.35,
    border: "1px solid #E5E7EB",
  },

  userBubble: {
    background: "#111827",
    color: "#FFFFFF",
    borderColor: "#111827",
  },

  assistantBubble: {
    background: "#FFFFFF",
    color: "#111827",
  },

  errorLine: {
    marginTop: 8,
    padding: "8px 10px",
    borderRadius: 12,
    background: "#FEF9C3",
    border: "1px solid #FDE68A",
    color: "#92400E",
    fontSize: 12,
  },

  footer: {
    flexShrink: 0,
    padding: 12,
    borderTop: "1px solid #E5E7EB",
    display: "flex",
    gap: 10,
    alignItems: "flex-end",
    background: "rgba(249,250,251,0.9)",
  },

  checkboxRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    padding: "6px 8px",
    border: "1px solid #E5E7EB",
    borderRadius: 12,
    background: "rgba(255,255,255,0.9)",
    fontSize: 12,
    color: "#111827",
    userSelect: "none",
    whiteSpace: "nowrap",
    cursor: "pointer",
  },
  checkbox: {
    width: 14,
    height: 14,
    cursor: "pointer",
  },
  checkboxLabel: { fontSize: 11, color: "#111827" },


  textarea: {
    flex: 1,
    background: "#ffffff",
    border: "2px solid #e5e7eb",
    resize: "none",
    borderRadius: 14,
    border: "1px solid #D1D5DB",
    padding: "10px 12px",
    fontSize: 12,
    outline: "none",
    background: "#FFFFFF",
  },

  sendBtn: {
    borderRadius: 14,
    border: "1px solid #0D9488",
    background: "#0D9488",
    color: "#FFFFFF",
    fontWeight: 900,
    padding: "10px 12px",
    cursor: "pointer",
    fontSize: 12,
  },

  smallBtn: {
  fontSize: 11,
  padding: "4px 8px",
  borderRadius: 8,
  border: "1px solid #e5e7eb",
  background: "#fff",
  cursor: "pointer",
  lineHeight: 1.2,
},

        queryBlock: {
  display: "flex",
  flexDirection: "column",
  gap: 8,
},

queryLabel: {
  fontSize: 12,
  fontWeight: 600,
},

textarea: {
  width: "100%",
  resize: "none",
  border: "1px solid #e5e7eb",
  borderRadius: 8,
  padding: 8,
  fontSize: 13,
},


  btnDisabled: {
    opacity: 0.55,
    cursor: "not-allowed",
  },
};
