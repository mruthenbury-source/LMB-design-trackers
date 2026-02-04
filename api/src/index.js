// api/src/index.js
// Never let one broken module prevent all routes from registering.

try {
  await import("./functions/bootstrap.js");
} catch (e) {
  console.error("bootstrap.js failed to load (continuing):", e);
}

import "./functions/save.js";
import "./functions/tick.js";
import "./functions/health.js";
import "./functions/backupWeekly.js";
import "./functions/chat.js";
import "./functions/exportPdf.js";
import "./functions/state.js";
import "./functions/meta.js";

