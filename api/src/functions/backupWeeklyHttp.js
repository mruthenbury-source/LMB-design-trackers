// api/src/functions/backupWeeklyHttp.js
// HTTP endpoint used by the client to discover weekly backup snapshots.
// The timer function `backupWeekly` creates/updates an index blob; this
// endpoint simply returns that index (or an empty array if none exists).

import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

function getContainer() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  if (!conn) throw new Error("Storage connection string not set (BLOB_CONNECTION_STRING or AzureWebJobsStorage)");
  const service = BlobServiceClient.fromConnectionString(conn);
  return service.getContainerClient(containerName);
}

app.http("backupWeekly", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (_req, context) => {
    try {
      const container = getContainer();
      await container.createIfNotExists();

      const backupPrefix = process.env.BACKUP_PREFIX || "backups/weekly/";
      const indexBlobName = process.env.BACKUP_INDEX_BLOB || `${backupPrefix}index.json`;
      const blob = container.getBlobClient(indexBlobName);

      const exists = await blob.exists();
      if (!exists) {
        // Return 200 with an empty payload so the client doesn't treat it as a hard error.
        return { status: 200, jsonBody: { ok: true, indexBlobName, backups: [] } };
      }

      const dl = await blob.download();
      const text = await streamToString(dl.readableStreamBody);
      let backups = [];
      try {
        const parsed = text ? JSON.parse(text) : [];
        backups = Array.isArray(parsed) ? parsed : [];
      } catch {
        backups = [];
      }

      return { status: 200, jsonBody: { ok: true, indexBlobName, backups } };
    } catch (e) {
      context?.log?.error?.(e);
      // Use 200 with ok:false to avoid breaking the UI flow if storage is not configured.
      return { status: 200, jsonBody: { ok: false, error: String(e), backups: [] } };
    }
  },
});
