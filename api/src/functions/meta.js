import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function safeStr(v){return (v??"").toString().trim();}

function getBlobClients() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  if (!conn) throw new Error("Storage connection string is not set (BLOB_CONNECTION_STRING or AzureWebJobsStorage)");
  const service = BlobServiceClient.fromConnectionString(conn);
  return { container: service.getContainerClient(containerName) };
}

async function streamToString(readable) {
  return await new Promise((resolve, reject) => {
    const chunks = [];
    readable.on("data", (d) => chunks.push(d));
    readable.on("end", () => resolve(Buffer.concat(chunks).toString("utf8")));
    readable.on("error", reject);
  });
}

async function readJsonBlob(container, blobName, maxBytes = 5_000_000) {
  const blob = container.getBlobClient(blobName);
  const props = await blob.getProperties().catch(() => null);
  if (props && props.contentLength && props.contentLength > maxBytes) {
    throw new Error(`Blob too large (${props.contentLength} bytes)`);
  }
  const dl = await blob.download();
  const text = await streamToString(dl.readableStreamBody);
  return text ? JSON.parse(text) : null;
}

app.http("meta", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async () => {
    try {
      const { container } = getBlobClients();
      const stateBlob = process.env.STATE_BLOB || "state.json";
      const state = await readJsonBlob(container, stateBlob);
      const projects = (state?.projects || []).map((p) => safeStr(p?.name)).filter(Boolean);
      return { status: 200, jsonBody: { ok: true, projects } };
    } catch (e) {
      return { status: 500, jsonBody: { ok: false, error: "meta handler error", details: String(e) } };
    }
  },
});
