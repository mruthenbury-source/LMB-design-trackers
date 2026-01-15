import { app } from "@azure/functions";
import { readAppState, appendBackup } from "../shared/sharepoint.js";

function isoDate(d = new Date()) {
  const yyyy = d.getUTCFullYear();
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  const dd = String(d.getUTCDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

// Every Sunday at 02:00 UTC
app.timer("backupWeekly", {
  schedule: "0 0 2 * * 0",
  handler: async (myTimer, context) => {
    try {
      const state = await readAppState();
      if (!state) return;
      await appendBackup(state, isoDate());
      context.log(`Backup written for ${isoDate()}`);
    } catch (e) {
      context.error(e);
    }
  },
});
