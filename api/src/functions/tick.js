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

function uid() {
  return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

function userLabel(me) {
  return (me && (me.name || me.email || me.userId)) || "Unknown";
}

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

      const isAdmin = roleNorm === "admin" || roleNorm === "owner";
      const now = new Date().toISOString();
      const by = userLabel(me);

      const updatedRow = { ...ref.rows[ref.idx] };
      updatedRow.locks = updatedRow.locks && typeof updatedRow.locks === "object" ? updatedRow.locks : {};
      updatedRow.auditTrail = Array.isArray(updatedRow.auditTrail) ? updatedRow.auditTrail : [];

      for (const k of keys) {
        const nextVal = !!patch[k];
        const locked = !!updatedRow.locks?.[k]?.locked;

        if (!nextVal && locked) {
          // Unlock request
          if (!isAdmin) return { status: 403, jsonBody: { error: "locked" } };
          updatedRow[k] = false;
          delete updatedRow.locks[k];
          updatedRow.auditTrail.unshift({ id: uid(), at: now, by, action: "unlock", field: k, value: false });
          continue;
        }

        updatedRow[k] = nextVal;

        if (nextVal) {
          // Lock on tick
          updatedRow.locks[k] = { locked: true, by, at: now };
          updatedRow.auditTrail.unshift({ id: uid(), at: now, by, action: "tick", field: k, value: true });
        } else {
          updatedRow.auditTrail.unshift({ id: uid(), at: now, by, action: "untick", field: k, value: false });
        }
      }

      // If Not Required is set, clear other checkmarks + remove locks
      if (updatedRow.notRequired) {
        for (const f of ["completed", "statusADone", "firstIssueDone"]) {
          if (updatedRow[f]) {
            updatedRow[f] = false;
            delete updatedRow.locks[f];
            updatedRow.auditTrail.unshift({ id: uid(), at: now, by, action: "auto_clear", field: f, value: false, note: "Not required" });
          } else {
            // still remove lock if present
            if (updatedRow.locks?.[f]?.locked) delete updatedRow.locks[f];
          }
        }
      }

      ref.rows[ref.idx] = updatedRow;

      await writeAppState(state);
      return { status: 200, jsonBody: { ok: true, row: updatedRow } };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: "tick_failed" } };
    }
  },
});
