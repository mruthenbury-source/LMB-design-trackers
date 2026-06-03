/**
 * POST /api/mobile/upload
 *
 * Handles file uploads from the mobile app. Builds the Azure Blob Storage folder
 * path dynamically from the project's schema (per-project configurable labels and
 * structure), then indexes the upload as a comment into the project state.
 *
 * Folder path pattern:
 *   {prefix}/{ProjectName}/{UploadType}/{DependentName}/{ContainerName}/{SubContainerName}/{file}
 *
 * All three hierarchy levels use the ACTUAL ITEM NAMES entered by the user in the
 * web app, so renaming "Block A" to "Tower A" automatically changes the folder path.
 * The number of path segments also adjusts: project-wide uploads omit all three
 * levels; package-wide uploads omit container/subContainer; exact row uploads
 * include all levels present for that row.
 */
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";
import crypto from "node:crypto";
import { getUserContext } from "../shared/userContext.js";

/* ── tiny helpers ── */
function s(v) { return (v ?? "").toString().trim(); }
function low(v) { return s(v).toLowerCase(); }
function today() { return new Date().toISOString().slice(0, 10); }
function uid(pfx = "mob") { return `${pfx}-${Date.now()}-${crypto.randomUUID?.() ?? crypto.randomBytes(8).toString("hex")}`; }

function sanitizeFileName(name = "file") {
  return (s(name) || "file").replace(/[^a-zA-Z0-9._-]+/g, "_").slice(0, 140);
}
function sanitizeSegment(name = "") {
  return (s(name) || "unknown").replace(/[\\/:*?"<>|#%{}~&]+/g, "-").replace(/\s+/g, " ").trim().slice(0, 120);
}

/* ── blob clients ── */
function getStateBlob() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const container = process.env.STATE_CONTAINER || "workback";
  const blob = process.env.STATE_BLOB_NAME || "state.json";
  if (!conn) throw new Error("Storage connection string not set");
  const svc = BlobServiceClient.fromConnectionString(conn);
  return svc.getContainerClient(container).getBlockBlobClient(blob);
}

function getUploadContainer() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const container = process.env.UPLOAD_CONTAINER || process.env.PROGRAMME_STATE_CONTAINER || "programme";
  const prefix = s(process.env.UPLOAD_PREFIX || process.env.PROGRAMME_COMMENT_ATTACHMENTS_PREFIX) || "documents";
  if (!conn) throw new Error("Storage connection string not set");
  const svc = BlobServiceClient.fromConnectionString(conn);
  return { client: svc.getContainerClient(container), prefix };
}

/* ── state helpers ── */
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

async function writeState(state) {
  const blob = getStateBlob();
  await blob.uploadData(Buffer.from(JSON.stringify(state)), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

/* ── schema helpers ── */
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

function resolveSchema(raw) {
  return { ...defaultSchema(), ...(raw || {}) };
}

/* ── upload type normalisation ── */
const UPLOAD_TYPE_MAP = new Map([
  ["drawing", "Drawings"], ["drawings", "Drawings"],
  ["specification", "Specifications"], ["specifications", "Specifications"],
  ["instruction", "Instructions"], ["instructions", "Instructions"],
  ["contract", "Contract"], ["contracts", "Contract"],
  ["image", "Images"], ["images", "Images"], ["photo", "Images"], ["photos", "Images"],
  ["report", "Reports"], ["reports", "Reports"],
  ["qa", "QA"], ["certificate", "Certificates"], ["certificates", "Certificates"],
]);
function normalizeUploadType(v) { return UPLOAD_TYPE_MAP.get(low(v)) || "Documents"; }

/* ── scope normalisation ── */
function normalizeScope(v) {
  const r = low(v);
  // Explicit values from the mobile app take priority
  if (r === "tracker" || r === "row") return "tracker";
  if (r === "package") return "package";
  if (r === "master") return "master";
  if (r === "project") return "project";
  // Legacy fallback: infer from string content
  if (r.includes("tracker") || r.includes("row") || r.includes("line")) return "tracker";
  if (r.includes("package")) return "package";
  if (r.includes("block") || r.includes("level") || r.includes("master")) return "master";
  return "project";
}

/* ── project lookup ── */
function projectIdOf(p, i) {
  return s(p.id || p.projectId || p.project_id || p.projectNumber || p.number || p.name || `project-${i}`);
}
function projectMatches(p, i, wantedId, wantedName) {
  const ids = [projectIdOf(p, i), p.id, p.projectId, p.name, p.projectName].map(low).filter(Boolean);
  return (wantedId && ids.includes(low(wantedId))) || (wantedName && ids.includes(low(wantedName)));
}
function findProject(state, projectId, projectName) {
  const projects = Array.isArray(state.projects) ? state.projects : [];
  const idx = projects.findIndex((p, i) => projectMatches(p, i, projectId, projectName));
  return idx >= 0 ? { project: projects[idx], index: idx, projects } : { project: null, index: -1, projects };
}

/* ── dependency/responsibility lookup ── */
function findDependent(project, packageName, responsibilityId) {
  const list = Array.isArray(project.responsibilities) ? project.responsibilities : [];
  return list.find((r) =>
    (responsibilityId && (s(r.id) === s(responsibilityId) || s(r.responsibilityId) === s(responsibilityId))) ||
    (packageName && (low(r.name) === low(packageName) || low(r.packageName) === low(packageName)))
  ) || null;
}

function supplierFromDependent(dep) {
  if (!dep) return "";
  return s(dep.supplier || dep.supplierName || dep.contractor || dep.company || dep.organization || "");
}

async function resolveSupplier({ projectId, projectName, packageName, responsibilityId, fallback }) {
  if (fallback && low(fallback) !== "mobile user") return fallback;
  try {
    const state = await readState();
    const { project } = findProject(state, projectId, projectName);
    if (!project) return fallback || "";
    const dep = findDependent(project, packageName, responsibilityId);
    return supplierFromDependent(dep) || fallback || "";
  } catch { return fallback || ""; }
}

/* ── folder path construction (per-project, schema-aware) ── */
/**
 * Builds the Azure Blob folder path for an upload.
 *
 * The path always starts with:
 *   {prefix}/{ProjectName}/{UploadType}/
 *
 * Then appends hierarchy segments depending on scope:
 *   project   → Project Home/
 *   package   → [supplier/]{dependentName}/
 *   master    → [supplier/]{containerName}/{subContainerName}/
 *   tracker   → [supplier/]{dependentName}/{containerName}/{subContainerName}/
 *
 * All segment names come from user-entered VALUES (not schema label names), so
 * renaming "Block A" to "Tower A" automatically changes the Azure folder path.
 * The supplier segment is omitted when blank.
 */
function buildFolderPath({ prefix, projectName, uploadType, scope, supplier, dependentName, containerName, subContainerName, fallbackPageName }) {
  const parts = [prefix, sanitizeSegment(projectName || "project"), normalizeUploadType(uploadType)];

  const addIf = (v) => { if (s(v)) parts.push(sanitizeSegment(v)); };
  const addSupplier = () => addIf(supplier);

  if (scope === "project") {
    parts.push("Project Home");
  } else if (scope === "package") {
    addSupplier();
    parts.push(sanitizeSegment(dependentName || "Package"));
  } else if (scope === "master") {
    addSupplier();
    parts.push(sanitizeSegment(fallbackPageName || "Project Home"));
    addIf(containerName);
    addIf(subContainerName);
  } else {
    // tracker — dependent + container + subContainer
    addSupplier();
    parts.push(sanitizeSegment(dependentName || "Package"));
    addIf(containerName);
    addIf(subContainerName);
  }

  return parts.join("/");
}

/* ── state indexing ── */
function buildScopeKey(scope, projectId, packageName, pageId, rowId, responsibilityId, suppliedKey) {
  if (s(suppliedKey).startsWith("project:")) return s(suppliedKey);
  const pid = projectId || "project";
  const pgid = pageId || "project-home";
  if (scope === "package") return `project:${pid}:page:${pgid}:package:${responsibilityId || packageName || "package"}`;
  if (scope === "project") return `project:${pid}:page:${pgid}:project-documents`;
  if (scope === "tracker") return `tracker:${pgid}:${rowId || "row"}`;
  return `master:${pgid}`;
}

function findTrackerRow(project, pageId, rowId, containerName, subContainerName, dependentName) {
  const pages = Array.isArray(project.pages) ? project.pages : [];
  for (const page of pages) {
    if (pageId && low(page.id || page.name) !== low(pageId)) continue;
    const rows = Array.isArray(page.rows) ? page.rows : [];
    for (const row of rows) {
      if (rowId && low(row.id || row.rowId) === low(rowId)) return { page, row };
      const rb = s(row.meta?.blockZone || row.blockZone || row.block || "");
      const rl = s(row.meta?.levelName || row.levelName || row.level || "");
      if (rb && rl && low(rb) === low(containerName) && low(rl) === low(subContainerName)) return { page, row };
    }
    // Synthesise a generated row if project uses generated rows from master
    if (containerName && subContainerName && !page.meta?.isMaster) {
      const master = Array.isArray(project.master) ? project.master : [];
      for (const m of master) {
        const bn = s(m.blockZone || m.block || m.zone || m.name);
        if (low(bn) !== low(containerName)) continue;
        for (const lv of Array.isArray(m.levels) ? m.levels : []) {
          const ln = s(lv.name || lv.levelName);
          if (low(ln) !== low(subContainerName)) continue;
          const genRow = {
            id: `generated::${page.name}::${bn}::${ln}`,
            kind: "item",
            item: `${page.name}_${bn}_${ln}`,
            completed: false, notRequired: false, statusADone: false, firstIssueDone: false,
            locks: {}, comments: [],
            meta: { generated: true, blockZone: bn, levelId: lv.id || null, levelName: ln, finishDate: lv.finishDate || "" },
            anchorKey: "requiredOnSite",
            anchorDateISO: lv.startDate || "",
          };
          rows.push(genRow);
          return { page, row: genRow };
        }
      }
    }
  }
  return { page: null, row: null };
}

function findMasterLevel(project, containerName, subContainerName) {
  for (const m of Array.isArray(project.master) ? project.master : []) {
    const bn = s(m.blockZone || m.block || m.zone || m.name);
    if (containerName && low(bn) !== low(containerName)) continue;
    for (const lv of Array.isArray(m.levels) ? m.levels : []) {
      const ln = s(lv.name || lv.levelName);
      if (subContainerName && low(ln) !== low(subContainerName)) continue;
      return { masterBlock: m, masterLevel: lv, containerName: bn, subContainerName: ln };
    }
  }
  return null;
}

async function indexUploadToState({ state, attachment, scope, dependentName, containerName, subContainerName, pageId, rowId, commentText, commentTags, userName, responsibilityId, suppliedScopeKey }) {
  const { project, index, projects } = findProject(state, attachment.projectId, attachment.projectName);
  if (!project) return { indexed: false, reason: "project_not_found" };

  const dep = findDependent(project, dependentName, responsibilityId);
  const resolvedResponsibilityId = s(dep?.id || dep?.responsibilityId || responsibilityId || "");
  const resolvedDependentName = s(dep?.name || dep?.packageName || dependentName || "");

  const scopeKey = buildScopeKey(scope, attachment.projectId, resolvedDependentName, pageId, rowId, resolvedResponsibilityId, suppliedScopeKey);

  const comment = {
    id: attachment.commentId || uid("mobile-comment"),
    name: userName || "Mobile User",
    dateISO: today(),
    createdAt: attachment.uploadedAt || new Date().toISOString(),
    text: commentText || `Uploaded ${attachment.fileName}`,
    tags: Array.isArray(commentTags) ? commentTags : [],
    pageId, pageName: attachment.pageName || "Project Home",
    uploadScopeKey: scopeKey,
    uploadFolderScope: resolvedDependentName || attachment.pageName || "Project Home",
    rowId, block: containerName, level: subContainerName,
    packageName: resolvedDependentName, item: resolvedDependentName, lineItemName: resolvedDependentName,
    responsibilityId: resolvedResponsibilityId,
    scopeType: scope,
    attachment: { ...attachment, uploadScopeKey: scopeKey, scopeType: scope, responsibilityId: resolvedResponsibilityId },
    locked: true, lockedBy: userName || "Mobile User",
    lockedAt: new Date().toISOString(),
    source: "tracksdn-mobile",
  };

  if (scope === "project") {
    (project.projectUploads = project.projectUploads || []).push(comment);
  } else if (scope === "package") {
    if (dep) {
      (dep.uploads = dep.uploads || []).push(comment);
    } else {
      (project.projectUploads = project.projectUploads || []).push({ ...comment, _warning: "dependent_not_found" });
    }
  } else if (scope === "tracker") {
    const { page, row } = findTrackerRow(project, pageId, rowId, containerName, subContainerName, dependentName);
    if (row) {
      (row.comments = row.comments || []).push(comment);
      comment.pageId = s(page?.id || pageId);
    } else {
      (project.projectUploads = project.projectUploads || []).push({ ...comment, _warning: "row_not_found" });
    }
  } else {
    // master
    const found = findMasterLevel(project, containerName, subContainerName);
    if (found) {
      (found.masterLevel.comments = found.masterLevel.comments || []).push(comment);
    } else {
      (project.projectUploads = project.projectUploads || []).push({ ...comment, _warning: "master_level_not_found" });
    }
  }

  projects[index] = project;
  await writeState({ ...state, projects });
  return { indexed: true, scope, scopeKey };
}

/* ── form parsing ── */
function getFiles(form) {
  const candidates = [];
  for (const key of ["files", "file", "attachments", "attachment"]) {
    try {
      const vals = typeof form.getAll === "function" ? form.getAll(key) : [form.get(key)];
      for (const v of vals || []) if (v && typeof v.arrayBuffer === "function") candidates.push(v);
    } catch {}
  }
  const seen = new Set();
  return candidates.filter((f) => {
    const sig = `${f.name}|${f.size}`;
    if (seen.has(sig)) return false;
    seen.add(sig);
    return true;
  });
}

function parseJsonField(raw, fallback = {}) {
  try { const p = JSON.parse(String(raw || "")); return p && typeof p === "object" && !Array.isArray(p) ? p : fallback; }
  catch { return fallback; }
}

function parseTags(raw) {
  try { const p = JSON.parse(String(raw || "[]")); return Array.isArray(p) ? p.map(s).filter(Boolean) : []; }
  catch { return []; }
}

function attachmentKind(contentType, fileName) {
  const t = low(contentType); const n = low(fileName);
  if (t === "application/pdf" || n.endsWith(".pdf")) return "pdf";
  if (t.startsWith("image/") || /\.(png|jpe?g|webp|gif|bmp|heic)$/i.test(n)) return "image";
  if (t.startsWith("video/")) return "video";
  return "file";
}

function safeMeta(obj) {
  const out = {};
  for (const [k, v] of Object.entries(obj || {})) {
    const key = String(k).toLowerCase().replace(/[^a-z0-9_]/g, "").slice(0, 60);
    if (!key) continue;
    out[key] = s(v).normalize("NFKD").replace(/[^\x20-\x7E]+/g, " ").replace(/[\r\n]+/g, " ").trim().slice(0, 1024);
  }
  return out;
}

/* ── handler ── */
const handler = async (req) => {
  try {
    const devBypass = String(process.env.MOBILE_DEV_BYPASS || "").toLowerCase() === "true";
    const user = getUserContext(req);
    if (!devBypass && !user?.authenticated) return { status: 401, jsonBody: { ok: false, error: "unauthorized" } };

    const form = await req.formData();
    const files = getFiles(form);
    if (!files.length) return { status: 400, jsonBody: { ok: false, error: "no_files_received" } };

    const ctx = parseJsonField(form.get("context"), {});

    const projectId = s(form.get("projectId") || ctx.projectId || "");
    const projectName = s(form.get("projectName") || ctx.projectName || projectId || "project");
    const scopeRaw = s(form.get("scopeType") || form.get("uploadScope") || ctx.scopeType || ctx.scope || "project");
    const scope = normalizeScope(scopeRaw);

    const dependentName = s(form.get("packageName") || ctx.packageName || ctx.package || "");
    const containerName = s(form.get("block") || ctx.block || "");
    const subContainerName = s(form.get("level") || ctx.level || "");
    const pageId = s(form.get("pageId") || ctx.pageId || "");
    const rowId = s(form.get("rowId") || ctx.rowId || "");
    const pageName = s(form.get("pageName") || ctx.pageName || "Project Home");
    const responsibilityId = s(form.get("responsibilityId") || ctx.responsibilityId || "");
    const uploadType = normalizeUploadType(form.get("uploadType") || ctx.uploadType || "Documents");
    const commentText = s(form.get("comment") || form.get("commentText") || ctx.comment || "");
    const commentTags = parseTags(form.get("tags") || form.get("commentTags")) || ["Mobile Upload", uploadType];
    const suppliedScopeKey = s(form.get("uploadScopeKey") || ctx.uploadScopeKey || "");

    let supplier = s(form.get("supplier") || ctx.supplier || "");
    supplier = await resolveSupplier({ projectId, projectName, packageName: dependentName, responsibilityId, fallback: supplier });

    if (scope === "tracker" && (!dependentName || !containerName || !subContainerName)) {
      return {
        status: 400,
        jsonBody: {
          ok: false,
          error: "missing_tracker_fields",
          message: "Tracker uploads require packageName, block, and level. Use /api/mobile/upload-targets to select a valid location.",
          received: { dependentName, containerName, subContainerName },
        },
      };
    }

    const { client: uploadClient, prefix } = getUploadContainer();
    await uploadClient.createIfNotExists();

    const folderPath = buildFolderPath({
      prefix, projectName, uploadType, scope, supplier,
      dependentName, containerName, subContainerName,
      fallbackPageName: pageName,
    });

    // Read state once for all files
    const state = await readState();

    const attachments = [];
    for (const file of files) {
      const rawName = sanitizeFileName(file.name || "attachment");
      const blobName = `${folderPath}/${Date.now()}-${uid("f")}-${rawName}`;
      const blobClient = uploadClient.getBlockBlobClient(blobName);
      const buf = Buffer.from(await file.arrayBuffer());
      const contentType = s(file.type) || "application/octet-stream";
      const kind = attachmentKind(contentType, rawName);
      const nowIso = new Date().toISOString();
      const attachmentId = uid("attachment");
      const commentId = uid("comment");

      await blobClient.uploadData(buf, {
        blobHTTPHeaders: { blobContentType: contentType },
        metadata: safeMeta({
          originalfilename: rawName,
          projectid: projectId,
          projectname: projectName,
          packagename: dependentName,
          block: containerName,
          level: subContainerName,
          supplier,
          uploadtype: uploadType,
          scope,
          pageid: pageId,
          pagename: pageName,
          mobileupload: "true",
          attachmentid: attachmentId,
          commentid: commentId,
        }),
      });

      const attachment = {
        id: attachmentId, attachmentId, commentId,
        fileName: rawName, blobName, blobPath: blobName,
        url: blobClient.url,
        size: Number(file.size) || buf.length,
        contentType, kind,
        uploadedAt: nowIso, uploadedBy: user?.name || user?.email || "Mobile User",
        projectId, projectName, pageId, pageName, rowId,
        supplier, dependentName, containerName, subContainerName,
        uploadType, scopeType: scope,
        folderPath,
        source: "tracksdn-mobile",
      };

      const indexResult = await indexUploadToState({
        state, attachment, scope,
        dependentName, containerName, subContainerName,
        pageId, rowId, commentText, commentTags,
        userName: user?.name || user?.email || "Mobile User",
        responsibilityId, suppliedScopeKey,
      });

      attachments.push({ ...attachment, stateIndexed: indexResult.indexed, indexResult });
    }

    return {
      status: 200,
      jsonBody: {
        ok: true,
        count: attachments.length,
        folderPath,
        attachment: attachments[0] || null,
        attachments,
      },
    };
  } catch (e) {
    return { status: 500, jsonBody: { ok: false, error: "upload_failed", details: String(e?.message || e) } };
  }
};

app.http("mobileUpload", {
  route: "mobile/upload",
  methods: ["GET", "POST", "OPTIONS"],
  authLevel: "anonymous",
  handler: async (req) => {
    if (req.method !== "POST") return { status: 200, jsonBody: { ok: true, route: "/api/mobile/upload", methods: ["POST"] } };
    return handler(req);
  },
});
