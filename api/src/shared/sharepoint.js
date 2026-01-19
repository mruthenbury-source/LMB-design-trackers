import { getGraphClient } from "./graphClient.js";

const listIdCache = new Map();

function env(name, fallback = "") {
  return process.env[name] || fallback;
}

export async function getListId(displayName) {
  const siteId = env("SP_SITE_ID");
  if (!siteId) throw new Error("Missing SP_SITE_ID");
  const key = `${siteId}::${displayName}`;
  if (listIdCache.has(key)) return listIdCache.get(key);

  const client = await getGraphClient();
  const lists = await client.api(`/sites/${siteId}/lists`).get();
  const found = (lists?.value || []).find((l) => String(l.displayName) === String(displayName));
  if (!found) throw new Error(`List not found: ${displayName}`);
  listIdCache.set(key, found.id);
  return found.id;
}

export async function findStateItemId(listName) {
  const siteId = env("SP_SITE_ID");
  const listId = await getListId(listName);
  const client = await getGraphClient();

  // Query items where fields/Title == 'STATE'
  // Graph supports $filter on fields in beta; for v1 we just fetch first page and scan.
  const items = await client.api(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`).get();
  const found = (items?.value || []).find((it) => (it?.fields?.Title || "") === "STATE");
  return found?.id || null;
}

export async function readAppState() {
  const siteId = env("SP_SITE_ID");
  const listName = env("SP_LIST_STATE", "WorkbackState");
  const listId = await getListId(listName);
  const client = await getGraphClient();

  const items = await client.api(`/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`).get();
  const found = (items?.value || []).find((it) => (it?.fields?.Title || "") === "STATE");
  if (!found?.fields?.DataJson) return null;

  try {
    return JSON.parse(found.fields.DataJson);
  } catch {
    return null;
  }
}

export async function writeAppState(state) {
  const siteId = env("SP_SITE_ID");
  const listName = env("SP_LIST_STATE", "WorkbackState");
  const listId = await getListId(listName);
  const client = await getGraphClient();

  const json = JSON.stringify(state);
  const now = new Date().toISOString();

  const itemId = await findStateItemId(listName);
  if (!itemId) {
    await client.api(`/sites/${siteId}/lists/${listId}/items`).post({
      fields: {
        Title: "STATE",
        DataJson: json,
        LastSavedUtc: now,
      },
    });
    return;
  }

  await client.api(`/sites/${siteId}/lists/${listId}/items/${itemId}/fields`).patch({
    Title: "STATE",
    DataJson: json,
    LastSavedUtc: now,
  });
}

export async function readPermissions() {
  const siteId = env("SP_SITE_ID");
  const listName = env("SP_LIST_PERMISSIONS", "WorkbackPermissions");
  const listId = await getListId(listName);
  const client = await getGraphClient();

  // Pull all permissions (keep small; split by project later)
  const out = [];
  let url = `/sites/${siteId}/lists/${listId}/items?$expand=fields&$top=200`;
  while (url) {
    const page = await client.api(url).get();
    (page?.value || []).forEach((it) => out.push(it.fields || {}));
    url = page["@odata.nextLink"]
      ? page["@odata.nextLink"].replace("https://graph.microsoft.com/v1.0", "")
      : null;
  }
  return out;
}

export async function readPermissionsForUser(user) {
  const rows = await readPermissions();
  const email = String(user?.email || "").toLowerCase();
  const oid = String(user?.oid || "").toLowerCase();

  // Build pageId -> role for this user
  const rolesByPageId = {};
  let strongest = "viewer";

  const rank = { viewer: 0, readOnly: 0, tickOnly: 1, checkbox: 1, editor: 2, admin: 3, owner: 3 };

  for (const r of rows) {
    const rEmail = String(r.UserEmail || "").toLowerCase();
    const rOid = String(r.UserObjectId || "").toLowerCase();
    const match = (oid && rOid && rOid === oid) || (!oid && email && rEmail && rEmail === email) || (email && rEmail && rEmail === email);
    if (!match) continue;

    const pageId = String(r.PageId || "");
    const role = String(r.Role || "viewer");
    if (pageId) rolesByPageId[pageId] = role;

    if ((rank[role] ?? 0) > (rank[strongest] ?? 0)) strongest = role;
  }

  return { role: strongest, rolesByPageId };
}

export async function appendBackup(state, suffixISODate) {
  const siteId = env("SP_SITE_ID");
  const listName = env("SP_LIST_BACKUPS", "WorkbackBackups");
  const listId = await getListId(listName);
  const client = await getGraphClient();

  const title = `backup-${suffixISODate}`;
  await client.api(`/sites/${siteId}/lists/${listId}/items`).post({
    fields: {
      Title: title,
      DataJson: JSON.stringify(state),
      CreatedUtc: new Date().toISOString(),
    },
  });
}
