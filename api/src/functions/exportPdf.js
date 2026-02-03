import { app } from "@azure/functions";
import PDFDocument from "pdfkit";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// ---- CONFIG YOU CAN EDIT ----
const BRAND = {
  companyName: "LMB Design Trackers",      // footer text
  // Dark grey header bar (avoid bright colours for print)
  headerFill: "#111827",
  headerText: "#FFFFFF",
  subText: "#6B7280",
  tableHeaderFill: "#F3F4F6",
  zebraFill: "#FAFAFA",
  line: "#D1D5DB",
  // Change to your file name if needed:
  logoFile: "logo.png",
};

// Keep answer text within page width (PDFKit does wrapping automatically if width is set)
function safe(v) {
  if (v === null || v === undefined) return "";
  return typeof v === "string" ? v : JSON.stringify(v);
}

function loadLogo() {
  try {
    return fs.readFileSync(path.join(__dirname, "..", "assets", BRAND.logoFile));
  } catch {
    return null;
  }
}

function drawHeaderFooter(doc, title, logoBuf, pageNo) {
  const L = doc.page.margins.left;
  const R = doc.page.width - doc.page.margins.right;
  const T = doc.page.margins.top;
  const B = doc.page.height - doc.page.margins.bottom;

  // HEADER BAR
  const barH = 42;
  doc.save();
  doc.rect(0, 0, doc.page.width, barH).fill(BRAND.headerFill);

  // Logo
  if (logoBuf) {
    // Place inside header bar with padding; size to fit bar height
    doc.image(logoBuf, L, 8, { height: barH - 16 });
  }

  // Title text (white)
  const titleX = L + (logoBuf ? 120 : 0);
  doc
    .fillColor(BRAND.headerText)
    .font("Helvetica-Bold")
    .fontSize(14)
    .text(title, titleX, 12, { width: R - titleX, align: "left" });

  doc.restore();

  // FOOTER
  doc.save();
  doc.strokeColor(BRAND.line).moveTo(L, B + 12).lineTo(R, B + 12).stroke();
  doc
    .fillColor(BRAND.subText)
    .font("Helvetica")
    .fontSize(9)
    .text(`${BRAND.companyName} • Generated ${new Date().toLocaleString()}`, L, B + 18, { align: "left" });

  doc
    .fillColor(BRAND.subText)
    .font("Helvetica")
    .fontSize(9)
    .text(`Page ${pageNo}`, L, B + 18, { width: R - L, align: "right" });

  doc.restore();
}

app.http("exportPdf", {
  methods: ["POST"],
  authLevel: "anonymous",

  handler: async (req) => {
    try {
      const body = await req.json().catch(() => ({}));
      const title = safe(body.title || "Chat Export");
      const answer = safe(body.answer || "");
      const rows = Array.isArray(body.rows) ? body.rows : [];
      const meta = body.meta && typeof body.meta === "object" ? body.meta : null;

      const doc = new PDFDocument({ size: "A4", margin: 40 });
      const buffers = [];
      doc.on("data", (d) => buffers.push(d));
      const done = new Promise((r) => doc.on("end", r));

      const logo = loadLogo();
      let pageNo = 1;

      // Draw on first page + every new page
      drawHeaderFooter(doc, title, logo, pageNo);
      doc.on("pageAdded", () => {
        pageNo++;
        drawHeaderFooter(doc, title, logo, pageNo);
      });

      // Start content below header bar
      doc.y = 60;

      // ---- PARAMETERS (optional) ----
      if (meta) {
        doc.font("Helvetica-Bold").fontSize(11).fillColor("#111827").text("Parameters");
        doc.moveDown(0.3);
        doc.font("Helvetica").fontSize(10).fillColor("#111827");

        Object.keys(meta).forEach((k) => {
          doc.fillColor(BRAND.subText).font("Helvetica-Bold").text(`${k}: `, { continued: true });
          doc.fillColor("#111827").font("Helvetica").text(safe(meta[k]));
        });

        doc.moveDown(0.8);
      }

      // ---- ANSWER TEXT ----
      doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Answer");
      doc.moveDown(0.5);
      doc.font("Helvetica").fontSize(10).fillColor("#111827");

      const pageW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      doc.text(answer, { width: pageW });

      // ---- TABLE ----
      if (rows.length) {
        doc.addPage();
        doc.y = 60;

        const pageW2 = doc.page.width - doc.page.margins.left - doc.page.margins.right;

        // Column widths (easy to tweak)
        // Make Item auto-fill remainder.
        const cols = [
          { key: "project", label: "Project", w: 85 },
          { key: "supplier", label: "Supplier", w: 70 },
          { key: "responsibility", label: "Resp", w: 80 },
          { key: "requiredOnSite", label: "Req", w: 55 },
          { key: "statusA", label: "A", w: 55 },
        ];
        const used = cols.reduce((a, c) => a + c.w, 0);
        const itemW = Math.max(140, pageW2 - used); // minimum width so it never collapses
        const allCols = [
          cols[0],
          cols[1],
          cols[2],
          { key: "item", label: "Item", w: itemW },
          cols[3],
          cols[4],
        ];

        const startX = doc.page.margins.left;
        let y = doc.y;
        const rowH = 16;
        const totalW = allCols.reduce((a, c) => a + c.w, 0);

        const ensureSpace = () => {
          if (y > doc.page.height - doc.page.margins.bottom - 60) {
            doc.addPage();
            doc.y = 60;
            y = doc.y;
            drawTableHeader(); // repeat header on new pages
          }
        };

        const drawTableHeader = () => {
          // Header background
          doc.save();
          doc.fillColor(BRAND.tableHeaderFill).rect(startX, y, totalW, rowH).fill();
          doc.restore();

          // Header text
          let x = startX;
          doc.font("Helvetica-Bold").fontSize(8).fillColor("#111827");
          for (const c of allCols) {
            doc.text(c.label, x + 2, y + 4, { width: c.w - 4, ellipsis: true });
            x += c.w;
          }

          // Header bottom line
          doc.save();
          doc.strokeColor(BRAND.line).moveTo(startX, y + rowH).lineTo(startX + totalW, y + rowH).stroke();
          doc.restore();

          y += rowH;
        };

        const drawRow = (r, idx) => {
          // Zebra
          if (idx % 2 === 1) {
            doc.save();
            doc.fillColor(BRAND.zebraFill).rect(startX, y, totalW, rowH).fill();
            doc.restore();
          }

          let x = startX;
          doc.font("Helvetica").fontSize(8).fillColor("#111827");
          for (const c of allCols) {
            doc.text(safe(r?.[c.key] ?? ""), x + 2, y + 4, {
              width: c.w - 4,
              ellipsis: true,
            });
            x += c.w;
          }

          // Row line
          doc.save();
          doc.strokeColor(BRAND.line).strokeOpacity(0.4).moveTo(startX, y + rowH).lineTo(startX + totalW, y + rowH).stroke();
          doc.restore();

          y += rowH;
          ensureSpace();
        };

        // Title
        doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Data");
        doc.moveDown(0.5);
        y = doc.y;

        drawTableHeader();

        const MAX = 1500; // bump if you want
        rows.slice(0, MAX).forEach((r, i) => drawRow(r, i));

        if (rows.length > MAX) {
          doc.moveDown(1);
          doc.font("Helvetica").fontSize(9).fillColor(BRAND.subText).text(`Showing first ${MAX} of ${rows.length} rows.`);
        }
      }

      doc.end();
      await done;

      return {
        status: 200,
        headers: {
          "Content-Type": "application/pdf",
          "Content-Disposition": 'attachment; filename="chat-export.pdf"',
        },
        body: Buffer.concat(buffers),
      };
    } catch (e) {
      return {
        status: 500,
        jsonBody: { error: "exportPdf failed", details: String(e) },
      };
    }
  },
});



