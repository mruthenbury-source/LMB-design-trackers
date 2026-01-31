
// src/ChatOverlay.jsx
import React, { useEffect, useMemo, useRef, useState } from "react";

/**
 * ChatOverlay
 * - Full-height right-side panel (h-screen)
 * - Header controls: Search backups (left) + Refresh + Close (small text)
 * - Keeps a free-typing input area for any query (always visible)
 * - Quick Questions dropdown inserts/sends structured prompts
 * - “All levels on site…” questions reference PROGRAMME data (not tracker)
 *
 * Props (adjust names if your app uses slightly different ones):
 * - isOpen: boolean
 * - onClose: () => void
 * - onRefresh?: () => void
 * - onSend: ({ message, searchBackups }) => Promise<void> | void
 * - chatMessages: Array<{ role: "user"|"assistant"|"system", content: string }>
 * - chatContext: object  (your existing combined context payload)
 * - programmeData?: object (optional if you pass programme separately)
 */
export default function ChatOverlay({
  isOpen,
  onClose,
  onRefresh,
  onSend,
  chatMessages = [],
  chatContext,
  programmeData,
}) {
  const [input, setInput] = useState("");
  const [searchBackups, setSearchBackups] = useState(false);

  // Quick question UI state
  const [quickId, setQuickId] = useState("");
  const [quickProject, setQuickProject] = useState("");
  const [quickOnSite, setQuickOnSite] = useState("");

  const listRef = useRef(null);

  // Auto-scroll to bottom when messages change
  useEffect(() => {
    if (!isOpen) return;
    const el = listRef.current;
    if (!el) return;
    el.scrollTop = el.scrollHeight;
  }, [isOpen, chatMessages.length]);

  const programmeHint = useMemo(() => {
    const hasProgramme =
      !!programmeData ||
      (chatContext && (chatContext.programme || chatContext.programmeData || chatContext.programmeIndex));
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
            `${programmeHint}`,
            `Using ONLY the PROGRAMME context, list ALL levels across ALL projects that are on site in: "${d}".`,
            "Interpret the input as a month/date/range and match levels scheduled on site within that period.",
            "Return: Project name, Level/Block, on-site date (or range), source field used.",
            "If programme dates are missing/ambiguous, say what’s missing and which fields you checked.",
          ].join("\n"),
      },
      {
        id: "onsite_project",
        label: "All levels on site in… — specific project (programme)",
        needsProject: true,
        needsOnSite: true,
        build: (p, d) =>
          [
            `${programmeHint}`,
            `Using ONLY the PROGRAMME context, list ALL levels for project: "${p}" that are on site in: "${d}".`,
            "Interpret the input as a month/date/range and match levels scheduled on site within that period.",
            "Return: Level/Block, on-site date (or range), source field used.",
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

    const p = quickProject.trim();
    const d = quickOnSite.trim();

    return selectedQuick.build(p, d);
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
    if (!text) return;

    await onSend?.({ message: text, searchBackups });
    setInput("");
  }

  if (!isOpen) return null;

  return (
    <div className="fixed top-0 right-0 h-screen w-[420px] max-w-[95vw] bg-white shadow-xl border-l flex flex-col z-50">
      <div className="flex items-center justify-between gap-2 px-3 py-2 border-b bg-white">
        <div className="font-semibold text-sm">Chat</div>

        <div className="flex items-center gap-3">
          <label className="flex items-center gap-1 text-xs select-none">
            <input
              type="checkbox"
              className="h-3 w-3"
              checked={searchBackups}
              onChange={(e) => setSearchBackups(e.target.checked)}
            />
            <span>Search backups</span>
          </label>

          <button
            type="button"
            onClick={() => onRefresh?.()}
            className="text-xs px-2 py-1 border rounded hover:bg-gray-50"
          >
            Refresh
          </button>

          <button
            type="button"
            onClick={onClose}
            className="text-xs px-2 py-1 border rounded hover:bg-gray-50"
          >
            Close
          </button>
        </div>
      </div>

      <div ref={listRef} className="flex-1 overflow-y-auto px-3 py-3 space-y-3">
        {chatMessages.map((m, idx) => {
          const isUser = m.role === "user";
          return (
            <div key={idx} className={`flex ${isUser ? "justify-end" : "justify-start"}`}>
              <div
                className={[
                  "max-w-[85%] rounded-lg px-3 py-2 text-sm whitespace-pre-wrap",
                  isUser ? "bg-gray-900 text-white" : "bg-gray-100 text-gray-900",
                ].join(" ")}
              >
                {m.content}
              </div>
            </div>
          );
        })}
        {chatMessages.length === 0 && (
          <div className="text-sm text-gray-500">
            Ask anything, or use a quick question below.
          </div>
        )}
      </div>

      <div className="border-t bg-white px-3 py-3 space-y-2">
        <div className="space-y-2">
          <div className="flex items-center gap-2">
            <select
              value={quickId}
              onChange={(e) => setQuickId(e.target.value)}
              className="flex-1 border rounded px-2 py-1 text-sm"
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
              className="text-xs px-2 py-1 border rounded hover:bg-gray-50"
              disabled={!buildQuickPrompt()}
              title="Insert into message box"
            >
              Insert
            </button>

            <button
              type="button"
              onClick={askQuick}
              className="text-xs px-2 py-1 border rounded hover:bg-gray-50"
              disabled={!buildQuickPrompt()}
              title="Send immediately"
            >
              Ask
            </button>
          </div>

          {(selectedQuick?.needsProject || selectedQuick?.needsOnSite) && (
            <div className="flex gap-2">
              {selectedQuick?.needsProject && (
                <input
                  value={quickProject}
                  onChange={(e) => setQuickProject(e.target.value)}
                  placeholder="Project name…"
                  className="flex-1 border rounded px-2 py-1 text-sm"
                />
              )}
              {selectedQuick?.needsOnSite && (
                <input
                  value={quickOnSite}
                  onChange={(e) => setQuickOnSite(e.target.value)}
                  placeholder="Month / date / range…"
                  className="flex-1 border rounded px-2 py-1 text-sm"
                />
              )}
            </div>
          )}
        </div>

        <div className="space-y-2">
          <textarea
            value={input}
            onChange={(e) => setInput(e.target.value)}
            placeholder="Ask any question…"
            rows={3}
            className="w-full border rounded p-2 text-sm resize-none"
          />

          <div className="flex justify-end">
            <button
              type="button"
              onClick={() => handleSend()}
              disabled={!input.trim()}
              className="text-sm px-3 py-2 bg-gray-900 text-white rounded disabled:opacity-50"
            >
              Send
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}
