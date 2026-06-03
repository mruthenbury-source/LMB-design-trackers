/**
 * GET  /api/hierarchy-config?projectId=xxx  → returns project schema + structure
 * PATCH /api/hierarchy-config?projectId=xxx  → updates project schema labels
 */
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function s(v) { return (v ?? "").toString().trim(); }
function first(...vals) { for (const v of vals) { const x = s(v); if (x) return x; } return ""; }

function getBlob() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  const blobName = process.env.STATE_BLOB_NAME || "state.json";
  if (!conn) throw new Error("Storage connection string not set");
  const svc = BlobServiceClient.fromConnectionString(conn);
  return svc.getContainerClient(containerName).getBlockBlobClient(blobName);
}

async function readRaw() {
  const blob = getBlob();
  if (!(await blob.exists())) return null;
  const dl = await blob.download();
  const chunks = [];
  await new Promise((res, rej) => {
    dl.readableStreamBody.on("data", (d) => chunks.push(d));
    dl.readableStreamBody.on("end", res);
    dl.readableStreamBody.on("error", rej);
  });
  return JSON.parse(Buffer.concat(chunks).toString("utf8"));
}

async function writeRaw(raw) {
  const blob = getBlob();
  await blob.uploadData(Buffer.from(JSON.stringify(raw)), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

function normalizeState(raw) {
  if (!raw) return { projects: [] };
  const state = raw?.state ?? raw;
  return Array.isArray(state) ? { projects: state } : (state || { projects: [] });
}

function projectIdOf(p, i) {
  return first(p.id, p.projectId, p.name, `project-${i}`);
}

function projectMatches(p, i, wanted) {
  const w = s(wanted).toLowerCase();
  if (!w) return false;
  return [projectIdOf(p, i), p.id, p.projectId, p.name]
    .map((x) => s(x).toLowerCase())
    .some((x) => x && x === w);
}

const ALLOWED_SCHEMA_KEYS = [
  "containerLabel", "containerPlural",
  "subContainerLabel", "subContainerPlural",
  "dependentLabel", "dependentPlural",
];

app.http("hierarchyConfig", {
  methods: ["GET", "PATCH"],
  authLevel: "anonymous",
  route: "hierarchy-config",
  handler: async (request, context) => {
    try {
      const url = new URL(request.url);
      const projectId = first(url.searchParams.get("projectId"), url.searchParams.get("id"));
      if (!projectId) return { status: 400, jsonBody: { ok: false, error: "projectId is required" } };

      const raw = await readRaw();
      const state = normalizeState(raw);
      const projects = Array.isArray(state.projects) ? state.projects : [];
      const idx = projects.findIndex((p, i) => projectMatches(p, i, projectId));

      if (idx < 0) return { status: 404, jsonBody: { ok: false, error: "project_not_found" } };

      const project = projects[idx];

      if (request.method === "GET") {
        return {
          status: 200,
          jsonBody: {
            ok: true,
            projectId: projectIdOf(project, idx),
            projectName: first(project.name, project.projectName),
            schema: project.schema || {},
          },
        };
      }

      // PATCH — update schema labels
      const body = await request.json().catch(() => ({}));
      const schemaPatch = body?.schema || {};
      const cleaned = {};
      for (const key of ALLOWED_SCHEMA_KEYS) {
        if (typeof schemaPatch[key] === "string" && schemaPatch[key].trim()) {
          cleaned[key] = schemaPatch[key].trim();
        }
      }

      projects[idx] = { ...project, schema: { ...(project.schema || {}), ...cleaned } };

      // Persist — preserve the wrapper if it existed
      const updated = raw?.state ? { ...raw, state: { ...state, projects } } : { ...state, projects };
      await writeRaw(updated);

      return {
        status: 200,
        jsonBody: { ok: true, schema: projects[idx].schema },
      };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { ok: false, error: String(e?.message || e) } };
    }
  },
});
