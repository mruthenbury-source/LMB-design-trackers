import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function getBlobClient() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  if (!conn) throw new Error("Storage connection string is not set (BLOB_CONNECTION_STRING or AzureWebJobsStorage)");
  const service = BlobServiceClient.fromConnectionString(conn);
  const container = service.getContainerClient(containerName);
  return { container };
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

function safeIsoFromBackupName(name) {
  // backups/state-YYYY-MM-DD.json
  const m = String(name || "").match(/state-(\d{4}-\d{2}-\d{2})\.json$/);
  return m ? m[1] : "";
}

function buildSearchableText(row) {
  const parts = [
    row?.projectName,
    row?.pageName,
    row?.supplier,
    row?.item,
    row?.anchorKey,
    row?.anchorDateISO,
    row?.requiredOnSite,
    row?.statusA,
    row?.firstIssue,
    row?.commentsText,
  ];
  return parts.filter(Boolean).join("\n").toLowerCase();
}

function buildBackupIndexFromState(state) {
  const projects = Array.isArray(state?.projects) ? state.projects : [];
  const g1 = Number.isFinite(state?.globalDaysReqToStatusA) ? state.globalDaysReqToStatusA : 0;
  const g2 = Number.isFinite(state?.globalDaysStatusAToFirstIssue) ? state.globalDaysStatusAToFirstIssue : 0;

  const parseISO = (value) => {
    if (!value) return null;
    const [y, m, d] = String(value).split("-").map(Number);
    if (!y || !m || !d) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCFullYear() !== y || dt.getUTCMonth() !== m - 1 || dt.getUTCDate() !== d) return null;
    return dt;
  };
  const formatISO = (dt) => {
    if (!dt) return "";
    const y = dt.getUTCFullYear();
    const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
    const d = String(dt.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  };
  const addDaysUTC = (dt, days) => {
    const out = new Date(dt.getTime());
    out.setUTCDate(out.getUTCDate() + days);
    return out;
  };
  const clampInt = (v, fallback = 0) => {
    const n = Number(v);
    if (!Number.isFinite(n)) return fallback;
    return Math.trunc(n);
  };

  const computeDates = ({ anchorKey, anchorDateISO, daysReqToStatusA, daysStatusAToFirstIssue }) => {
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
  };

  const rows = [];
  projects.forEach((proj) => {
    const supplierByRespId = new Map((proj.responsibilities || []).map((r) => [r.id, r.supplier || ""]));

    (proj.pages || []).forEach((pg) => {
      const supplier = supplierByRespId.get(pg.meta?.responsibilityId) || "";

      (pg.rows || []).forEach((r) => {
        const d1 = r.overrideDaysReqToStatusA ?? g1;
        const d2 = r.overrideDaysStatusAToFirstIssue ?? g2;
        const dates = computeDates({
          anchorKey: r.anchorKey,
          anchorDateISO: r.anchorDateISO,
          daysReqToStatusA: d1,
          daysStatusAToFirstIssue: d2,
        });

        const comments = Array.isArray(r.comments) ? r.comments : [];
        const commentsText = comments
          .map((c) => {
            const who = String(c?.name || "").trim();
            const when = String(c?.at || "").trim();
            const txt = String(c?.text || "").trim();
            const head = [who, when].filter(Boolean).join(" ");
            return head + (txt ? `: ${txt}` : "");
          })
          .filter(Boolean)
          .join("\n");

        rows.push({
          projectName: proj.name,
          pageName: pg.name,
          supplier,
          item: r.item,
          anchorKey: r.anchorKey,
          anchorDateISO: r.anchorDateISO,
          requiredOnSite: dates.requiredOnSite,
          statusA: dates.statusA,
          firstIssue: dates.firstIssue,
          commentCount: comments.length,
          commentsText,
        });
      });
    });
  });

  return rows;
}

async function searchBackups({ queryText, maxBackups = 12, maxMatches = 60 }) {
  const q = String(queryText || "").trim().toLowerCase();
  if (!q) return { backupsScanned: 0, matches: [] };

  const { container } = getBlobClient();
  await container.createIfNotExists();

  const prefix = process.env.BACKUP_PREFIX || "backups/";
  const blobs = [];
  for await (const b of container.listBlobsFlat({ prefix })) {
    blobs.push({
      name: b.name,
      lastModified: b.properties?.lastModified ? new Date(b.properties.lastModified) : null,
      iso: safeIsoFromBackupName(b.name),
    });
  }

  blobs.sort((a, b) => (b.lastModified?.getTime() ?? 0) - (a.lastModified?.getTime() ?? 0));
  const scan = blobs.slice(0, maxBackups);

  const matches = [];
  for (const meta of scan) {
    if (matches.length >= maxMatches) break;

    const blob = container.getBlobClient(meta.name);
    const dl = await blob.download();
    const text = await streamToString(dl.readableStreamBody);

    let state = null;
    try {
      state = text ? JSON.parse(text) : null;
    } catch {
      state = null;
    }
    if (!state) continue;

    const rows = buildBackupIndexFromState(state);
    for (const row of rows) {
      if (matches.length >= maxMatches) break;
      const hay = buildSearchableText(row);
      if (hay.includes(q)) {
        matches.push({
          backup: meta.iso || meta.name,
          project: row.projectName,
          responsibility: row.pageName,
          supplier: row.supplier || "",
          item: row.item,
          anchorDateISO: row.anchorDateISO,
          requiredOnSite: row.requiredOnSite,
          statusA: row.statusA,
          firstIssue: row.firstIssue,
          commentCount: row.commentCount,
          commentsText: row.commentsText,
        });
      }
    }
  }

  return { backupsScanned: scan.length, matches };
}

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const { messages, context: appContext, searchBackups: wantBackups } = await req.json();

      // 🔒 Server-side safety: block supplier/guest users even if they try to call the endpoint directly.
      if (appContext?.user?.role === "guest" || appContext?.user?.isGuest === true) {
        return {
          status: 403,
          jsonBody: { error: "Chat is disabled for guest users" },
        };
      }

      const apiKey = process.env.OPENAI_API_KEY;
      const model = process.env.OPENAI_MODEL || "gpt-4o-mini";

      if (!apiKey) {
        return {
          status: 500,
          jsonBody: {
            error:
              "OPENAI_API_KEY is not set. Add it in Azure Static Web Apps > Configuration > Application settings.",
          },
        };
      }

      const systemText =
        "You are a helpful assistant for a design programme tracker web app. " +
        "Use the provided APP_CONTEXT_JSON to answer questions accurately. " +
        "If something is missing, say what you’d need. " +
        "If BACKUP_SEARCH_RESULTS is provided, treat it as historic snapshot data (may not match current state).";

      let backupContext = null;
      if (wantBackups) {
        const lastUser = Array.isArray(messages)
          ? [...messages].reverse().find((m) => m?.role === "user" && String(m?.content || "").trim())
          : null;
        const queryText = String(lastUser?.content || "");

        try {
          const r = await searchBackups({ queryText });
          backupContext = {
            enabled: true,
            query: queryText,
            backupsScanned: r.backupsScanned,
            matchCount: r.matches.length,
            matches: r.matches,
          };
        } catch (e) {
          backupContext = { enabled: true, error: String(e?.message || e) };
        }
      }

      const input = [
        { role: "system", content: systemText },
        {
          role: "system",
          content: `APP_CONTEXT_JSON:\n${JSON.stringify(
            {
              ...(appContext ?? {}),
              ...(backupContext ? { BACKUP_SEARCH_RESULTS: backupContext } : {}),
            },
            null,
            2
          )}`,
        },
        ...(Array.isArray(messages) ? messages : []),
      ];

      const r = await fetch("https://api.openai.com/v1/responses", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: `Bearer ${apiKey}`,
        },
        body: JSON.stringify({
          model,
          input,
          temperature: 0.2,
          max_output_tokens: 800,
        }),
      });

      const data = await r.json().catch(() => ({}));
      if (!r.ok) {
        return {
          status: 500,
          jsonBody: {
            error: "OpenAI request failed",
            status: r.status,
            details: data,
          },
        };
      }

      const answer = data?.output_text || data?.output?.[0]?.content?.[0]?.text || "No response text returned.";

      return {
        status: 200,
        jsonBody: { answer },
      };
    } catch (e) {
      return {
        status: 500,
        jsonBody: { error: "chat handler error", details: String(e) },
      };
    }
  },
});
