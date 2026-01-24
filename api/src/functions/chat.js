import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function getContainer() {
  const conn =
    process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  if (!conn) throw new Error("Storage connection string not set");
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(
    containerName
  );
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () =>
      resolve(Buffer.concat(chunks).toString("utf8"))
    );
    readable.on("error", reject);
  });
}

async function loadWeeklyBackups(limit = 3) {
  const container = getContainer();
  const prefix = process.env.BACKUP_PREFIX || "backups/";
  const found = [];

  for await (const blob of container.listBlobsFlat({ prefix })) {
    found.push({
      name: blob.name,
      lastModified: blob.properties?.lastModified,
    });
  }

  found.sort(
    (a, b) => new Date(b.lastModified) - new Date(a.lastModified)
  );

  const picked = found.slice(0, limit);
  const backups = [];

  for (const b of picked) {
    const bb = container.getBlockBlobClient(b.name);
    const dl = await bb.download();
    const text = await streamToString(dl.readableStreamBody);
    backups.push({
      date: b.lastModified,
      state: text ? JSON.parse(text) : null,
    });
  }

  return backups;
}

app.http("chat", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const { messages, context: appContext, searchBackups } =
        await req.json();

      const apiKey = process.env.OPENAI_API_KEY;
      const model = process.env.OPENAI_MODEL || "gpt-4o-mini";

      if (!apiKey) {
        return {
          status: 500,
          jsonBody: { error: "OPENAI_API_KEY is not set" },
        };
      }

      let backups = [];
      if (searchBackups) {
        try {
          backups = await loadWeeklyBackups(3);
        } catch {
          backups = [];
        }
      }

      const mergedContext = {
        ...appContext,
        backupsEnabled: !!searchBackups,
        backups,
      };

      const systemText =
        "You are a helpful assistant for a design programme tracker web app called SupplySync.\n" +
        "Use APP_CONTEXT_JSON to answer questions.\n" +
        "Important rules:\n" +
        "- Tick boxes are explicit booleans (statusADone, firstIssueDone, completed, notRequired).\n" +
        "- Dates may include requiredOnSite, statusA, firstIssue, programme start/finish.\n" +
        "- Durations may be provided in days or derived from dates.\n" +
        "- If backupsEnabled=true, compare current state with backups when asked about history.\n" +
        "- If something is missing, say what is required.";

      const input = [
        { role: "system", content: systemText },
        {
          role: "system",
          content: `APP_CONTEXT_JSON:\n${JSON.stringify(
            mergedContext,
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
          max_output_tokens: 900,
        }),
      });

      const data = await r.json().catch(() => ({}));

      if (!r.ok) {
        return {
          status: 500,
          jsonBody: { error: "OpenAI request failed", details: data },
        };
      }

      const answer =
        data?.output_text ||
        data?.output?.[0]?.content?.[0]?.text ||
        "No response text returned.";

      return { status: 200, jsonBody: { answer } };
    } catch (e) {
      return {
        status: 500,
        jsonBody: { error: "chat handler error", details: String(e) },
      };
    }
  },
});
