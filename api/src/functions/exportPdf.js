import { app } from "@azure/functions";
import PDFDocument from "pdfkit";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const BRAND = {
  companyName: "Lee Marley Brickwork",
  headerFill: "#111827",
  headerText: "#FFFFFF",
  subText: "#6B7280",
  tableHeaderFill: "#F3F4F6",
  zebraFill: "#FAFAFA",
  line: "#D1D5DB",
  logoFile: "logo.png",
};

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
  const B = doc.page.height - doc.page.margins.bottom;

  // Header bar (fixed top)
  const barH = 42;
  doc.save();
  doc.rect(0, 0, doc.page.width, barH).fill(BRAND.headerFill);

  if (logoBuf) doc.image(logoBuf, L, 8, { height: barH - 16 });

  const titleX = L + (logoBuf ? 120 : 0);
  doc
    .fillColor(BRAND.headerText)
    .font("Helvetica-Bold")
    .fontSize(14)
    .text(title, titleX, 12, { width: R - titleX, lineBreak: false });

  doc.restore();

  // Footer (keep inside page)
  const footerY = B + 10; // inside bottom margin area
  doc.save();
  doc.strokeColor(BRAND.line).moveTo(L, footerY).lineTo(R, footerY).stroke();

  doc.font("Helvetica").fontSize(9).fillColor(BRAND.subText);
  doc.text(`${BRAND.companyName} • Generated ${new Date().toLocaleString()}`, L, footerY + 6, {
    width: R - L,
    align: "left",
    lineBreak: false,
  });
  doc.text(`Page ${pageNo}`, L, footerY + 6, {
    width: R - L,
    align: "right",
    lineBreak: false,
  });
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

      const addPageWithChrome = () => {
        doc.addPage();
        pageNo += 1;
        drawHeaderFooter(doc, title, logo, pageNo);
        doc.y = 60; // below header bar
      };

      // First page chrome
      drawHeaderFooter(doc, title, logo, pageNo);
      doc.y = 60;

      // ---- Parameters (optional) ----
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

      // ---- Answer ----
      doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Answer");
      doc.moveDown(0.5);
      doc.font("Helvetica").fontSize(10).fillColor("#111827");

      const pageW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      doc.text(answer, { width: pageW });

      // ---- Table ----
      if (rows.length) {
        addPageWithChrome();

        const pageW2 = doc.page.width - doc.page.margins.left - doc.page.margins.right;

        // Column widths (easy to tweak)
        const baseCols = [
          { key: "project", label: "Project", w: 85 },
          { key: "supplier", label: "Supplier", w: 70 },
          { key: "responsibility", label: "Resp", w: 80 },
          { key: "requiredOnSite", label: "Req", w: 55 },
          { key: "statusA", label: "A", w: 55 },
        ];
        const used = baseCols.reduce((a, c) => a + c.w, 0);
        const itemW = Math.max(160, pageW2 - used);

        const cols = [
          baseCols[0],
          baseCols[1],
          baseCols[2],
          { key: "item", label: "Item", w: itemW },
          baseCols[3],
          baseCols[4],
        ];

        const startX = doc.page.margins.left;
        let y = doc.y;
        const rowH = 16;
        const totalW = cols.reduce((a, c) => a + c.w, 0);

        const ensureSpace = () => {
          const bottomLimit = doc.page.height - doc.page.margins.bottom - 60;
          if (y > bottomLimit) {
            addPageWithChrome();
            y = doc.y;
            drawHeaderRow();
          }
        };

        const drawHeaderRow = () => {
          doc.save();
          doc.fillColor(BRAND.tableHeaderFill).rect(startX, y, totalW, rowH).fill();
          doc.restore();

          let x = startX;
          doc.font("Helvetica-Bold").fontSize(8).fillColor("#111827");
          cols.forEach((c) => {
            doc.text(c.label, x + 2, y + 4, { width: c.w - 4, lineBreak: false, ellipsis: true });
            x += c.w;
          });

          doc.save();
          doc.strokeColor(BRAND.line).moveTo(startX, y + rowH).lineTo(startX + totalW, y + rowH).stroke();
          doc.restore();

          y += rowH;
        };

        const drawDataRow = (r, idx) => {
          if (idx % 2 === 1) {
            doc.save();
            doc.fillColor(BRAND.zebraFill).rect(startX, y, totalW, rowH).fill();
            doc.restore();
          }

          let x = startX;
          doc.font("Helvetica").fontSize(8).fillColor("#111827");
          cols.forEach((c) => {
            doc.text(safe(r?.[c.key] ?? ""), x + 2, y + 4, { width: c.w - 4, lineBreak: false, ellipsis: true });
            x += c.w;
          });

          doc.save();
          doc.strokeColor(BRAND.line).strokeOpacity(0.35).moveTo(startX, y + rowH).lineTo(startX + totalW, y + rowH).stroke();
          doc.restore();

          y += rowH;
          ensureSpace();
        };

        doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Data");
        doc.moveDown(0.5);
        y = doc.y;

        drawHeaderRow();

        const MAX = 1500;
        rows.slice(0, MAX).forEach((r, i) => drawDataRow(r, i));

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
      return { status: 500, jsonBody: { error: "exportPdf failed", details: String(e) } };
    }
  },
});




