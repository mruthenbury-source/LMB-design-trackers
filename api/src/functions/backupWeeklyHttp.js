// api/src/functions/backupWeeklyHttp.js
// HTTP-triggered weekly backup. Intended for Static Web Apps (Free tier) where timers won't run.
// Call this endpoint from the client when the app is open (e.g., Mondays after 9am UK time).

import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function pad2(n) {
  return String(n).padStart(2, "0");
}

function ymdUTC(d = new Date()) {
  const y = d.getUTCFullYear();
  const m = pad2(d.getUTCMonth() + 1);
  const day = pad2(d.getUTCDate());
  return `${y}-${m}-${day}`;
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

app.http("backupWeekly", {
  methods: ["POST"],
  authLevel: "anonymous", // SWA handles auth/roles; route config enforces "authenticated"
  handler: async (req, context) => {
    try {
      const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
      const containerName = process.env.STATE_CONTAINER || "workback";
      const stateBlobName = process.env.STATE_BLOB_NAME || "state.json";

      const backupPrefix = process.env.BACKUP_PREFIX || "backups/weekly/";
      const indexBlobName = process.env.BACKUP_INDEX_BLOB || `${backupPrefix}index.json`;

      if (!conn) {
        return { status: 500, jsonBody: { ok: false, error: "Storage connection string not set" } };
      }

      // Optional: client can pass a deterministic weekKey like "2026-02-02"
      // so the backup name is stable and idempotent across devices.
      let weekKey = null;
      try {
        const body = await req.json();
        weekKey = body?.weekKey || null;
      } catch {
        // no body
      }
      if (!weekKey) weekKey = ymdUTC(new Date());

      const service = BlobServiceClient.fromConnectionString(conn);
      const container = service.getContainerClient(containerName);
      await container.createIfNotExists();

      // Ensure source exists
      const src = container.getBlobClient(stateBlobName);
      const srcExists = await src.exists();
      if (!srcExists) {
        return { status: 200, jsonBody: { ok: true, skipped: true, reason: "No state blob yet", source: stateBlobName } };
      }

      // Deterministic destination name: one per weekKey
      const backupName = `${backupPrefix}state-${weekKey}.json`;
      const dest = container.getBlobClient(backupName);

      // Idempotency: if backup already exists, return without changing it
      if (await dest.exists()) {
        return { status: 200, jsonBody: { ok: true, already: true, name: backupName } };
      }

      context.log(`[backupWeekly] Copying ${stateBlobName} -> ${backupName}`);
      const poller = await dest.beginCopyFromURL(src.url);
      const copyResult = await poller.pollUntilDone();

      // Update index.json
      const indexBlob = container.getBlockBlobClient(indexBlobName);
      let index = [];
      try {
        if (await indexBlob.exists()) {
          const dl = await indexBlob.download();
          const txt = await streamToString(dl.readableStreamBody);
          index = txt ? JSON.parse(txt) : [];
          if (!Array.isArray(index)) index = [];
        }
      } catch {
        index = [];
      }

      let srcProps = null;
      try {
        srcProps = await src.getProperties();
      } catch {}

      index.unshift({
        name: backupName,
        createdAtUTC: new Date().toISOString(),
        weekKey,
        sourceBlob: stateBlobName,
        sourceETag: srcProps?.etag || null,
        copyId: copyResult?.copyId || null,
      });
      index = index.slice(0, 200);

      await indexBlob.uploadData(Buffer.from(JSON.stringify(index, null, 2)), {
        blobHTTPHeaders: { blobContentType: "application/json" },
      });

      return { status: 200, jsonBody: { ok: true, name: backupName } };
    } catch (e) {
      context.log(`[backupWeekly] ERROR: ${e?.message || e}`);
      return { status: 500, jsonBody: { ok: false, error: String(e?.message || e) } };
    }
  },
});
