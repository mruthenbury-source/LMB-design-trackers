import { app } from "@azure/functions";
import { getUserContext } from "../shared/userContext.js";
import { readAppState, writeAppState, readPermissionsForUser } from "../shared/sharepoint.js";

function findRowRef(state, rowId) {
  for (const proj of state.projects || []) {
    for (const pg of proj.pages || []) {
      const rows = pg.rows || [];
      for (let i = 0; i < rows.length; i++) {
        if (rows[i].id === rowId) return { proj, pg, rows, idx: i };
      }
    }
  }
  return null;
}

const ALLOWED_FIELDS = ["completed", "notRequired", "statusADone", "firstIssueDone"];

app.http("rowTick", {
  route: "rows/{rowId}/tick",
  methods: ["PATCH"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      const rowId = request.params.rowId;
      const me = getUserContext(request);
      const perms = me.authenticated ? await readPermissionsForUser(me).catch(() => ({ rolesByPageId: {}, role: "viewer" })) : { rolesByPageId: {}, role: "viewer" };

      const body = (await request.json().catch(() => null)) || {};
      const patch = body.patch || {};

      // Basic validation
      const keys = Object.keys(patch);
      if (!keys.length || !keys.every((k) => ALLOWED_FIELDS.includes(k))) {
        return { status: 400, jsonBody: { error: "invalid_patch" } };
      }

      const state = await readAppState();
      const ref = findRowRef(state, rowId);
      if (!ref) return { status: 404, jsonBody: { error: "row_not_found" } };

      // Permission check is page-based
      const pageRole = perms.rolesByPageId?.[ref.pg.id] || perms.role || "viewer";
      const roleNorm = String(pageRole).toLowerCase();
      const ok = roleNorm === "tickonly" || roleNorm === "checkbox" || roleNorm === "editor" || roleNorm === "admin" || roleNorm === "owner";
      if (!ok) return { status: 403, jsonBody: { error: "forbidden" } };

      const updatedRow = { ...ref.rows[ref.idx] };
      for (const k of keys) updatedRow[k] = !!patch[k];

      // If Not Required is set, clear other checkmarks
      if (updatedRow.notRequired) {
        updatedRow.completed = false;
        updatedRow.statusADone = false;
        updatedRow.firstIssueDone = false;
      }

      ref.rows[ref.idx] = updatedRow;

      await writeAppState(state);
      return { status: 200, jsonBody: { ok: true } };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: "tick_failed" } };
    }
  },
});
