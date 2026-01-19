import { app } from "@azure/functions";
import { getUserContext } from "../shared/userContext.js";
import { readAppState, readPermissionsForUser } from "../shared/sharepoint.js";

app.http("bootstrap", {
  methods: ["GET"],
  authLevel: "anonymous",
  handler: async (request, context) => {
    try {
      const me = getUserContext(request);
      const state = await readAppState().catch(() => null);
      const perms = me.authenticated ? await readPermissionsForUser(me).catch(() => ({ rolesByPageId: {}, role: "viewer" })) : { rolesByPageId: {}, role: "viewer" };

      if (!state) {
        return {
          status: 200,
          jsonBody: {
            projects: [],
            settings: null,
            view: "landing",
            summaryFilter: "ongoing",
            summaryProjectId: "all",
            summarySupplier: "all",
            activeProjectId: null,
            activePageId: null,
            me: { ...me, role: perms.role },
            rolesByPageId: perms.rolesByPageId,
          },
        };
      }

      return {
        status: 200,
        jsonBody: {
          ...state,
          me: { ...me, role: perms.role },
          rolesByPageId: perms.rolesByPageId,
        },
      };
    } catch (e) {
      context.error(e);
      return { status: 500, jsonBody: { error: "bootstrap_failed" } };
    }
  },
});
