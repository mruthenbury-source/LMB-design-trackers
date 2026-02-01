// api/src/functions/backupWeekly.js
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function pad2(n) {
  return String(n).padStart(2, "0");
}

function stampUTC(d = new Date()) {
  // 2026-01-26T0200Z (no ":" so it’s blob-name + URL friendly)
  const y = d.getUTCFullYear();
  const m = pad2(d.getUTCMonth() + 1);
  const day = pad2(d.getUTCDate());
  const hh = pad2(d.getUTCHours());
  const mm = pad2(d.getUTCMinutes());
  return `${y}-${m}-${day}T${hh}${mm}Z`;
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

app.timer("backupWeekly", {
  // Every Monday at 09:00 UTC (NCRONTAB format) :contentReference[oaicite:1]{index=1}
  schedule: "0 0 9 * * 1",
  handler: async (myTimer, context) => {
    const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
    const containerName = process.env.STATE_CONTAINER || "workback";
    const stateBlobName = process.env.STATE_BLOB_NAME || "state.json";

    // Where backups go
    const backupPrefix = process.env.BACKUP_PREFIX || "backups/weekly/";
    const indexBlobName = process.env.BACKUP_INDEX_BLOB || `${backupPrefix}index.json`;

    if (!conn) throw new Error("Storage connection string not set (BLOB_CONNECTION_STRING or AzureWebJobsStorage)");

    const service = BlobServiceClient.fromConnectionString(conn);
    const container = service.getContainerClient(containerName);
    await container.createIfNotExists();

    // Source (latest live state)
    const src = container.getBlobClient(stateBlobName);
    const exists = await src.exists();
    if (!exists) {
      context.log("[backupWeekly] No state blob yet - skipping backup.");
      return;
    }

    // Destination (weekly snapshot)
    const now = new Date();
    const stamp = stampUTC(now);
    const backupName = `${backupPrefix}state-${stamp}.json`;
    const dest = container.getBlobClient(backupName);

    // Copy live state -> backup
    context.log(`[backupWeekly] Copying ${stateBlobName} -> ${backupName}`);
    const poller = await dest.beginCopyFromURL(src.url);
    const copyResult = await poller.pollUntilDone();

    // Update index.json so chat can list/compare snapshots later
    const indexBlob = container.getBlockBlobClient(indexBlobName);
    let index = [];
    try {
      const hasIndex = await indexBlob.exists();
      if (hasIndex) {
        const dl = await indexBlob.download();
        const txt = await streamToString(dl.readableStreamBody);
        index = txt ? JSON.parse(txt) : [];
        if (!Array.isArray(index)) index = [];
      }
    } catch {
      index = [];
    }

    // Attempt to capture some metadata
    let srcProps = null;
    try {
      srcProps = await src.getProperties();
    } catch {}

    index.unshift({
      name: backupName,
      createdAtUTC: now.toISOString(),
      stamp,
      sourceBlob: stateBlobName,
      sourceETag: srcProps?.etag || null,
      copyId: copyResult?.copyId || null,
    });

    // keep last 200 entries
    index = index.slice(0, 200);

    await indexBlob.uploadData(Buffer.from(JSON.stringify(index, null, 2)), {
      blobHTTPHeaders: { blobContentType: "application/json" },
    });

    context.log(`[backupWeekly] Done. Index updated: ${indexBlobName}`);
  },
});


