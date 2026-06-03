import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";

function s(v) { return (v ?? "").toString().trim(); }
function first(...vals) { for (const v of vals) { const x = s(v); if (x) return x; } return ""; }

function getStateBlob() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const containerName = process.env.STATE_CONTAINER || "workback";
  const blobName = process.env.STATE_BLOB_NAME || "state.json";
  if (!conn) throw new Error("Storage connection string not set");
  const svc = BlobServiceClient.fromConnectionString(conn);
  return svc.getContainerClient(containerName).getBlockBlobClient(blobName);
}

async function readState() {
  const blob = getStateBlob();
  if (!(await blob.exists())) return { projects: [] };
  const dl = await blob.download();
  const chunks = [];
  await new Promise((res, rej) => {
    dl.readableStreamBody.on("data", (d) => chunks.push(d));
    dl.readableStreamBody.on("end", res);
    dl.readableStreamBody.on("error", rej);
  });
  const raw = JSON.parse(Buffer.concat(chunks).toString("utf8"));
  const state = raw?.state ?? raw;
  return Array.isArray(state) ? { projects: state } : (state || { projects: [] });
}

function defaultSchema() {
  return {
    containerLabel: "Block",
    containerPlural: "Blocks",
    subContainerLabel: "Level",
    subContainerPlural: "Levels",
    dependentLabel: "Package",
    dependentPlural: "Packages",
  };
}

function projectIdOf(p, i) {
  return first(p.id, p.projectId, p.project_id, p.projectNumber, p.number, p.name, `project-${i}`);
}

function projectMatches(p, i, wanted) {
  const w = s(wanted).toLowerCase();
  if (!w) return false;
  return [projectIdOf(p, i), p.id, p.projectId, p.name, p.projectName]
    .map((x) => s(x).toLowerCase())
    .some((x) => x && x === w);
}

function collectTargets(project) {
  const master = Array.isArray(project.master) ? project.master : [];
  const responsibilities = Array.isArray(project.responsibilities) ? project.responsibilities : [];
  const pages = Array.isArray(project.pages) ? project.pages : [];

  const containers = master.map((m) => ({
    id: first(m.id, m.blockId),
    name: first(m.blockZone, m.block, m.zone, m.name),
    subContainers: (Array.isArray(m.levels) ? m.levels : []).map((lv) => ({
      id: first(lv.id, lv.levelId),
      name: first(lv.name, lv.levelName),
      startDate: first(lv.startDate),
      finishDate: first(lv.finishDate),
    })),
  })).filter((c) => c.name);

  const dependents = responsibilities.map((r) => ({
    id: first(r.id, r.responsibilityId),
    name: first(r.name, r.packageName),
    supplier: first(r.supplier, r.supplierName),
    permissions: r.permissions || { users: [] },
  })).filter((d) => d.name);

  // Build flat rows: dependent × container × subContainer
  const rows = [];
  const rowKeys = new Set();
  for (const pg of pages) {
    if (pg.meta?.isMaster) continue;
    const pageId = first(pg.id, pg.pageId, pg.name);
    const pageName = first(pg.name, pg.pageName);
    const respId = first(pg.meta?.responsibilityId, pg.responsibilityId);
    const resp = responsibilities.find((r) => first(r.id, r.responsibilityId) === respId);
    const packageName = first(pg.meta?.packageName, pg.packageName, resp?.name, pageName);
    const supplierName = first(pg.supplierName, pg.supplier, resp?.supplier, resp?.supplierName);

    for (const row of Array.isArray(pg.rows) ? pg.rows : []) {
      const rowId = first(row.id, row.rowId);
      const block = first(row.meta?.blockZone, row.blockZone, row.block, row.Block, row.blockName);
      const level = first(row.meta?.levelName, row.levelName, row.level, row.Level);
      if (!pageId || !rowId || !packageName || !block || !level) continue;
      const key = `${pageId}|${rowId}`;
      if (rowKeys.has(key)) continue;
      rowKeys.add(key);
      rows.push({ pageId, pageName, rowId, packageName, supplierName, block, level });
    }

    // Synthetic rows from master block/levels for pages that have none stored
    if ((Array.isArray(pg.rows) ? pg.rows : []).length === 0) {
      for (const m of master) {
        const block = first(m.blockZone, m.block, m.zone, m.name);
        if (!block) continue;
        for (const lv of Array.isArray(m.levels) ? m.levels : []) {
          const level = first(lv.name, lv.levelName);
          if (!level) continue;
          const rowId = `generated::${pageName}::${block}::${level}`;
          const key = `${pageId}|${rowId}`;
          if (rowKeys.has(key)) continue;
          rowKeys.add(key);
          rows.push({ pageId, pageName, rowId, packageName, supplierName, block, level, generated: true });
        }
      }
    }
  }

  return { containers, dependents, rows };
}

app.http("mobileUploadTargets", {
  methods: ["GET"],
  authLevel: "anonymous",
  route: "mobile/upload-targets",
  handler: async (request, context) => {
    try {
      const url = new URL(request.url);
      const projectId = first(url.searchParams.get("projectId"), url.searchParams.get("id"));

      const state = await readState();
      const projects = Array.isArray(state.projects) ? state.projects : [];
      const idx = projects.findIndex((p, i) => projectMatches(p, i, projectId));

      if (idx < 0) {
        return {
          status: 404,
          jsonBody: {
            ok: false,
            error: "project_not_found",
            availableProjectIds: projects.map((p, i) => projectIdOf(p, i)).slice(0, 50),
          },
        };
      }

      const project = projects[idx];
      const schema = { ...defaultSchema(), ...(project.schema || {}) };
      const targets = collectTargets(project);

      return {
        status: 200,
        headers: { "Cache-Control": "no-store" },
        jsonBody: {
          ok: true,
          project: {
            id: projectIdOf(project, idx),
            name: first(project.name, project.projectName, "Untitled"),
          },
          schema,
          ...targets,
          counts: {
            containers: targets.containers.length,
            dependents: targets.dependents.length,
            rows: targets.rows.length,
          },
        },
      };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { ok: false, error: String(e?.message || e) } };
    }
  },
});
