import { app } from "@azure/functions";
import { getUserContext } from "../shared/userContext.js";
import { writeAppState, readPermissionsForUser } from "../shared/sharepoint.js";

app.http("save", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      const me = getUserContext(request);
      const perms = me.authenticated ? await readPermissionsForUser(me).catch(() => ({ role: "viewer" })) : { role: "viewer" };
      const role = String(perms.role || "viewer");

      const canWrite = role === "admin" || role === "owner" || role === "editor";
      if (!canWrite) {
        return { status: 403, jsonBody: { error: "forbidden", role } };
      }

      const body = await request.json();
      await writeAppState(body);
      return { status: 200, jsonBody: { ok: true } };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: "save_failed" } };
    }
  },
});
