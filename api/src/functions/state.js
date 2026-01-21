import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function getClients() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  const blobName = process.env.STATE_BLOB_NAME || "state.json";

  if (!conn) throw new Error("Storage connection string is not set (BLOB_CONNECTION_STRING)");

  const service = BlobServiceClient.fromConnectionString(conn);
  const container = service.getContainerClient(containerName);
  const blockBlob = container.getBlockBlobClient(blobName);

  return { container, blockBlob };
}
function getClientPrincipal(req) {
  const header = req.headers.get("x-ms-client-principal");
  if (!header) return null;
  try {
    const decoded = Buffer.from(header, "base64").toString("utf8");
    return JSON.parse(decoded);
  } catch {
    return null;
  }
}

function hasRole(principal, role) {
  const roles = principal?.userRoles || [];
  return roles.includes(role);
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
      const { container, blockBlob } = getClients();
      await container.createIfNotExists();

      if (req.method === "GET") {
        const exists = await blockBlob.exists();
        if (!exists) return { status: 200, jsonBody: { ok: true, state: null } };

        const dl = await blockBlob.download();
        const text = await streamToString(dl.readableStreamBody);
        const state = text ? JSON.parse(text) : null;

        return { status: 200, jsonBody: { ok: true, state } };
      }
      // POST: restrict writes
      if (req.method === "POST") {
        const principal = getClientPrincipal(req);

        // must be logged in AND be admin
        if (!principal) {
          return { status: 401, jsonBody: { ok: false, error: "Login required to save" } };
        }
        if (!hasRole(principal, "admin")) {
          return { status: 403, jsonBody: { ok: false, error: "Admin role required to save" } };
        }
      }

      // POST
      const body = await req.json();
      const state = body?.state ?? body;
      const json = JSON.stringify(state ?? null);

      await blockBlob.uploadData(Buffer.from(json), {
        blobHTTPHeaders: { blobContentType: "application/json" },
      });

      return { status: 200, jsonBody: { ok: true } };
    } catch (err) {
      return { status: 500, jsonBody: { ok: false, error: String(err?.message || err) } };
    }
  },
});
