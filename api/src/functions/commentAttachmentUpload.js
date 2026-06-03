/**
 * POST /api/commentAttachmentUpload
 *
 * Handles file uploads from the web app (drag-drop or file picker on a tracker row or
 * Project Home). Builds the Azure Blob folder path using the project's per-project
 * schema so the folder names in Azure match whatever the admin has configured.
 *
 * Folder path: {prefix}/{ProjectName}/{UploadType}/[supplier/]{dependentName}/{containerName}/{subContainerName}/
 */
import { app } from "@azure/functions";
import { BlobServiceClient } from "@azure/storage-blob";
import crypto from "node:crypto";
import { getUserContext } from "../shared/userContext.js";

function s(v) { return (v ?? "").toString().trim(); }
function low(v) { return s(v).toLowerCase(); }
function uid(pfx = "att") { return `${pfx}-${Date.now()}-${crypto.randomUUID?.() ?? crypto.randomBytes(8).toString("hex")}`; }

function sanitizeFileName(name = "file") {
  return (s(name) || "file").replace(/[^a-zA-Z0-9._-]+/g, "_").slice(0, 140);
}
function sanitizeSegment(name = "") {
  return (s(name) || "unknown").replace(/[\\/:*?"<>|#%{}~&]+/g, "-").replace(/\s+/g, " ").trim().slice(0, 120);
}

const UPLOAD_TYPE_MAP = new Map([
  ["drawing", "Drawings"], ["drawings", "Drawings"],
  ["specification", "Specifications"], ["specifications", "Specifications"],
  ["instruction", "Instructions"], ["instructions", "Instructions"],
  ["contract", "Contract"], ["image", "Images"], ["images", "Images"],
  ["photo", "Images"], ["report", "Reports"], ["reports", "Reports"],
  ["qa", "QA"], ["certificate", "Certificates"],
]);
function normalizeUploadType(v) { return UPLOAD_TYPE_MAP.get(low(v)) || "Documents"; }

function normalizeScope(v) {
  const r = low(v);
  if (r === "tracker" || r === "row") return "tracker";
  if (r === "package") return "package";
  if (r === "master") return "master";
  if (r === "project") return "project";
  if (r.includes("tracker") || r.includes("row")) return "tracker";
  if (r.includes("package")) return "package";
  if (r.includes("block") || r.includes("level") || r.includes("master")) return "master";
  return "project";
}

function getStateBlob() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const container = process.env.STATE_CONTAINER || "workback";
  const blob = process.env.STATE_BLOB_NAME || "state.json";
  if (!conn) throw new Error("Storage connection string not set");
  return BlobServiceClient.fromConnectionString(conn).getContainerClient(container).getBlockBlobClient(blob);
}

function getUploadContainer() {
  const conn = process.env.BLOB_CONNECTION_STRING || process.env.AzureWebJobsStorage;
  const container = process.env.UPLOAD_CONTAINER || process.env.PROGRAMME_STATE_CONTAINER || "programme";
  const prefix = s(process.env.UPLOAD_PREFIX || process.env.PROGRAMME_COMMENT_ATTACHMENTS_PREFIX) || "documents";
  if (!conn) throw new Error("Storage connection string not set");
  return { client: BlobServiceClient.fromConnectionString(conn).getContainerClient(container), prefix };
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

async function writeState(state) {
  const blob = getStateBlob();
  await blob.uploadData(Buffer.from(JSON.stringify(state)), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

function projectIdOf(p, i) {
  return s(p.id || p.projectId || p.project_id || p.name || `project-${i}`);
}
function findProject(state, projectId, projectName) {
  const projects = Array.isArray(state.projects) ? state.projects : [];
  const idx = projects.findIndex((p, i) => {
    const ids = [projectIdOf(p, i), p.id, p.projectId, p.name].map(low).filter(Boolean);
    return (projectId && ids.includes(low(projectId))) || (projectName && ids.includes(low(projectName)));
  });
  return idx >= 0 ? { project: projects[idx], index: idx, projects } : { project: null, index: -1, projects };
}

function findDependent(project, dependentName, responsibilityId) {
  const list = Array.isArray(project.responsibilities) ? project.responsibilities : [];
  return list.find((r) =>
    (responsibilityId && s(r.id) === s(responsibilityId)) ||
    (dependentName && low(r.name) === low(dependentName))
  ) || null;
}

/**
 * Build folder path — same logic as mobileUpload.js so web and mobile uploads
 * land in the same folder structure.
 */
function buildFolderPath({ prefix, projectName, uploadType, scope, supplier, dependentName, containerName, subContainerName, fallbackPageName }) {
  const parts = [prefix, sanitizeSegment(projectName || "project"), normalizeUploadType(uploadType)];
  const addIf = (v) => { if (s(v)) parts.push(sanitizeSegment(v)); };

  if (scope === "project") {
    parts.push("Project Home");
  } else if (scope === "package") {
    addIf(supplier);
    parts.push(sanitizeSegment(dependentName || "Package"));
  } else if (scope === "master") {
    addIf(supplier);
    parts.push(sanitizeSegment(fallbackPageName || "Project Home"));
    addIf(containerName);
    addIf(subContainerName);
  } else {
    // tracker
    addIf(supplier);
    parts.push(sanitizeSegment(dependentName || "Package"));
    addIf(containerName);
    addIf(subContainerName);
  }

  return parts.join("/");
}

async function indexCommentToState({ state, projectId, projectName, attachment, scope, dependentName, containerName, subContainerName, pageId, rowId, responsibilityId, commentText, tags, userName, suppliedScopeKey }) {
  const { project, index, projects } = findProject(state, projectId, projectName);
  if (!project) return;

  const dep = findDependent(project, dependentName, responsibilityId);
  const resolvedId = s(dep?.id || dep?.responsibilityId || responsibilityId || "");
  const resolvedName = s(dep?.name || dependentName || "");

  const buildKey = () => {
    if (s(suppliedScopeKey).startsWith("project:")) return s(suppliedScopeKey);
    const pid = projectId || "project";
    const pgid = pageId || "project-home";
    if (scope === "package") return `project:${pid}:page:${pgid}:package:${resolvedId || resolvedName || "package"}`;
    if (scope === "project") return `project:${pid}:page:${pgid}:project-documents`;
    if (scope === "tracker") return `tracker:${pgid}:${rowId || "row"}`;
    return `master:${pgid}`;
  };

  const scopeKey = buildKey();
  const comment = {
    id: uid("comment"),
    name: userName || "Web User",
    dateISO: new Date().toISOString().slice(0, 10),
    createdAt: new Date().toISOString(),
    text: commentText || `Uploaded ${attachment.fileName}`,
    tags: Array.isArray(tags) ? tags : [],
    pageId, pageName: attachment.pageName || "Project Home",
    uploadScopeKey: scopeKey,
    uploadFolderScope: resolvedName || attachment.pageName || "Project Home",
    rowId, block: containerName, level: subContainerName,
    packageName: resolvedName, item: resolvedName, lineItemName: resolvedName,
    responsibilityId: resolvedId, scopeType: scope,
    attachment: { ...attachment, uploadScopeKey: scopeKey, scopeType: scope },
    locked: false,
    source: "web",
  };

  if (scope === "project") {
    (project.projectUploads = project.projectUploads || []).push(comment);
  } else if (scope === "package" && dep) {
    (dep.uploads = dep.uploads || []).push(comment);
  } else if (scope === "tracker" || scope === "master") {
    const pages = Array.isArray(project.pages) ? project.pages : [];
    for (const page of pages) {
      if (pageId && low(page.id || page.name) !== low(pageId)) continue;
      const rows = Array.isArray(page.rows) ? page.rows : [];
      const row = rows.find((r) => rowId && low(r.id || r.rowId) === low(rowId));
      if (row) { (row.comments = row.comments || []).push(comment); break; }
    }
  } else {
    (project.projectUploads = project.projectUploads || []).push(comment);
  }

  projects[index] = project;
  await writeState({ ...state, projects });
}

function getFiles(form) {
  const seen = new Set();
  const out = [];
  for (const key of ["file", "files", "attachment", "attachments"]) {
    try {
      const vals = typeof form.getAll === "function" ? form.getAll(key) : [form.get(key)];
      for (const v of vals || []) {
        if (!v || typeof v.arrayBuffer !== "function") continue;
        const sig = `${v.name}|${v.size}`;
        if (seen.has(sig)) continue;
        seen.add(sig);
        out.push(v);
      }
    } catch {}
  }
  return out;
}

function attachmentKind(contentType, fileName) {
  const t = low(contentType); const n = low(fileName);
  if (t === "application/pdf" || n.endsWith(".pdf")) return "pdf";
  if (t.startsWith("image/") || /\.(png|jpe?g|webp|gif|bmp|heic)$/i.test(n)) return "image";
  return "file";
}

app.http("commentAttachmentUpload", {
  methods: ["POST"],
  authLevel: "anonymous",
  route: "commentAttachmentUpload",
  handler: async (request, context) => {
    try {
      const user = getUserContext(request);

      const form = await request.formData();
      const files = getFiles(form);
      if (!files.length) return { status: 400, jsonBody: { ok: false, error: "no_file" } };

      const projectId = s(form.get("projectId") || "");
      const projectName = s(form.get("projectName") || projectId || "project");
      const scopeRaw = s(form.get("scopeType") || form.get("uploadScope") || "project");
      const scope = normalizeScope(scopeRaw);
      const dependentName = s(form.get("packageName") || form.get("dependentName") || "");
      const containerName = s(form.get("block") || form.get("containerName") || "");
      const subContainerName = s(form.get("level") || form.get("subContainerName") || "");
      const pageId = s(form.get("pageId") || "");
      const rowId = s(form.get("rowId") || "");
      const pageName = s(form.get("pageName") || "Project Home");
      const responsibilityId = s(form.get("responsibilityId") || "");
      const supplier = s(form.get("supplier") || "");
      const uploadType = normalizeUploadType(form.get("uploadType") || "Documents");
      const commentText = s(form.get("comment") || form.get("commentText") || "");
      const tags = (() => { try { return JSON.parse(form.get("tags") || "[]"); } catch { return []; } })();
      const suppliedScopeKey = s(form.get("uploadScopeKey") || "");

      const { client: uploadClient, prefix } = getUploadContainer();
      await uploadClient.createIfNotExists();

      const folderPath = buildFolderPath({
        prefix, projectName, uploadType, scope, supplier,
        dependentName, containerName, subContainerName,
        fallbackPageName: pageName,
      });

      const state = await readState();
      const attachments = [];

      for (const file of files) {
        const rawName = sanitizeFileName(file.name || "attachment");
        const blobName = `${folderPath}/${Date.now()}-${uid()}-${rawName}`;
        const blobClient = uploadClient.getBlockBlobClient(blobName);
        const buf = Buffer.from(await file.arrayBuffer());
        const contentType = s(file.type) || "application/octet-stream";
        const kind = attachmentKind(contentType, rawName);
        const nowIso = new Date().toISOString();
        const attachmentId = uid("attachment");

        await blobClient.uploadData(buf, { blobHTTPHeaders: { blobContentType: contentType } });

        const attachment = {
          id: attachmentId, attachmentId,
          fileName: rawName, blobName, blobPath: blobName,
          url: blobClient.url,
          size: Number(file.size) || buf.length,
          contentType, kind,
          uploadedAt: nowIso,
          uploadedBy: user?.name || user?.email || "Web User",
          projectId, projectName, pageId, pageName, rowId,
          supplier, dependentName, containerName, subContainerName,
          uploadType, scopeType: scope, folderPath,
          source: "web",
        };

        await indexCommentToState({
          state, projectId, projectName, attachment, scope,
          dependentName, containerName, subContainerName,
          pageId, rowId, responsibilityId, commentText, tags,
          userName: user?.name || user?.email || "Web User",
          suppliedScopeKey,
        });

        attachments.push(attachment);
      }

      return {
        status: 200,
        jsonBody: { ok: true, count: attachments.length, folderPath, attachment: attachments[0] || null, attachments },
      };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { ok: false, error: "upload_failed", details: String(e?.message || e) } };
    }
  },
});
