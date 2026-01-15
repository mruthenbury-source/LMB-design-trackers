import React, { useEffect } from "react";

export default function ChatOverlay({
  open,
  setOpen,
  messages,
  input,
  setInput,
  busy,
  sendChat,
  endRef,
  status, // "idle" | "checking" | "ok" | "error"
  onReset, // ✅ NEW: function passed in from App
}) {
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
                sendChat?.();
              }
            }}
          />

          <button
            onClick={() => sendChat?.()}
            disabled={busy || !input.trim()}
            style={{
              ...styles.sendBtn,
              ...((busy || !input.trim()) ? styles.btnDisabled : null),
            }}
            title="Send (Enter)"
          >
            Send
          </button>
        </div>
      </div>
    </div>
  );
}

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
    inset: 0,
    zIndex: 60,
    background: "rgba(17,24,39,0.35)",
    display: "flex",
    alignItems: "flex-end",
    justifyContent: "flex-end",
    padding: 14,
  },

  panel: {
    width: "min(440px, 96vw)",
    height: "min(70vh, 720px)",
    background: "rgba(255,255,255,0.92)",
    border: "1px solid #E5E7EB",
    borderRadius: 18,
    boxShadow: "0 18px 60px rgba(17,24,39,0.25)",
    overflow: "hidden",
    display: "flex",
    flexDirection: "column",
    backdropFilter: "blur(10px)",
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
    padding: 12,
    borderTop: "1px solid #E5E7EB",
    display: "flex",
    gap: 10,
    alignItems: "flex-end",
    background: "rgba(249,250,251,0.9)",
  },

  textarea: {
    flex: 1,
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
    borderRadius: 12,
    border: "1px solid #D1D5DB",
    background: "#FFFFFF",
    color: "#111827",
    fontWeight: 800,
    padding: "8px 10px",
    cursor: "pointer",
    fontSize: 12,
  },

  btnDisabled: {
    opacity: 0.55,
    cursor: "not-allowed",
  },
};
