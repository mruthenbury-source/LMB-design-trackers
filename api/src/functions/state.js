// api/src/functions/state.js
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function getBlobClients() {
  const conn = process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  const blobName = process.env.STATE_BLOB_NAME || "state.json";

  if (!conn) throw new Error("AzureWebJobsStorage is not set");

  const service = BlobServiceClient.fromConnectionString(conn);
  const container = service.getContainerClient(containerName);
  const blob = container.getBlobClient(blobName);

  return { container, blob };
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

app.http("state", {
  methods: ["GET", "POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const { container, blob } = getBlobClients();
      await container.createIfNotExists();

      if (req.method === "GET") {
        const exists = await blob.exists();
        if (!exists) {
          return { status: 200, jsonBody: { ok: true, state: null } };
        }

        const dl = await blob.download();
        const text = await streamToString(dl.readableStreamBody);
        const state = text ? JSON.parse(text) : null;

        return { status: 200, jsonBody: { ok: true, state } };
      }

      // POST
      const body = await req.json();
      const state = body?.state ?? body;

      const json = JSON.stringify(state ?? null);
      await blob.upload(json, Buffer.byteLength(json), {
        blobHTTPHeaders: { blobContentType: "application/json" },
      });

      return { status: 200, jsonBody: { ok: true } };
    } catch (err) {
      return { status: 500, jsonBody: { ok: false, error: String(err?.message || err) } };
    }
  },
});
