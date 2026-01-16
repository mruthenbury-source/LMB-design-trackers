// api/src/functions/backupWeekly.js
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function isoDateUTC() {
  const d = new Date();
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

app.timer("backupWeekly", {
  // Every Monday at 02:00 UTC
  schedule: "0 0 2 * * 1",
  handler: async () => {
    const conn = process.env.AzureWebJobsStorage;
    const containerName = process.env.STATE_CONTAINER || "workback";
    const blobName = process.env.STATE_BLOB_NAME || "state.json";
    const prefix = process.env.BACKUP_PREFIX || "backups/";

    if (!conn) throw new Error("AzureWebJobsStorage is not set");

    const service = BlobServiceClient.fromConnectionString(conn);
    const container = service.getContainerClient(containerName);
    await container.createIfNotExists();

    const src = container.getBlobClient(blobName);
    const exists = await src.exists();
    if (!exists) return; // nothing to backup yet

    const backupName = `${prefix}state-${isoDateUTC()}.json`;
    const dest = container.getBlobClient(backupName);

    const poller = await dest.beginCopyFromURL(src.url);
    await poller.pollUntilDone();
  },
});
