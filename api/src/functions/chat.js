import { app } from "@azure/functions";
import PDFDocument from "pdfkit";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const BRAND = {
  companyName: "LMB Design Trackers",
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

function drawChrome(doc, title, logoBuf, pageNo) {
  const L = doc.page.margins.left;
  const R = doc.page.width - doc.page.margins.right;
  const B = doc.page.height - doc.page.margins.bottom;

  // Header bar
  const barH = 42;
  doc.save();
  doc.fillColor(BRAND.headerFill);
  doc.rect(0, 0, doc.page.width, barH).fill();

  if (logoBuf) doc.image(logoBuf, L, 8, { height: barH - 16 });

  const titleX = L + (logoBuf ? 120 : 0);
  doc.fillColor(BRAND.headerText).font("Helvetica-Bold").fontSize(14);
  doc.text(title, titleX, 12, { width: R - titleX, lineBreak: false });

  doc.restore();

  // Footer
  const footerY = B + 10;
  doc.save();
  doc.strokeColor(BRAND.line);
  doc.moveTo(L, footerY).lineTo(R, footerY).stroke();

  doc.fillColor(BRAND.subText).font("Helvetica").fontSize(9);
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

function buildColumnsFromRows(rows, pageW) {
  // Preferred tracker columns (your main table shape)
  const preferred = [
    { key: "project", label: "Project", w: 90 },
    { key: "supplier", label: "Supplier", w: 75 },
    { key: "responsibility", label: "Resp", w: 85 },
    { key: "item", label: "Item", w: 999 }, // auto-fill
    { key: "requiredOnSite", label: "Req", w: 65 },
    { key: "statusA", label: "A", w: 60 },
  ];

  const first = rows && rows.length ? rows[0] : null;
  const keys = first && typeof first === "object" ? Object.keys(first) : [];

  // If rows don’t match the preferred keys, fall back to detected keys
  const preferredKeys = new Set(preferred.map((c) => c.key));
  const hasPreferred = keys.some((k) => preferredKeys.has(k));

  let cols;
  if (hasPreferred) {
    cols = preferred;
  } else {
    // Use up to 6 keys from the data
    const chosen = keys.slice(0, 6);
    cols = chosen.map((k, idx) => ({
      key: k,
      label: k,
      w: idx === chosen.length - 1 ? 999 : Math.floor(pageW / Math.max(2, chosen.length)), // last auto-fill
    }));
    // ensure there is an auto-fill column
    if (cols.length) cols[cols.length - 1].w = 999;
  }

  // Auto-fit last/auto column to fill remaining width exactly
  const fixedW = cols.filter((c) => c.w !== 999).reduce((a, c) => a + c.w, 0);
  const autoIdx = cols.findIndex((c) => c.w === 999);
  const autoW = Math.max(160, pageW - fixedW);
  if (autoIdx >= 0) cols[autoIdx] = { ...cols[autoIdx], w: autoW };

  // Scale widths to fit exactly (prevents drift)
  const totalTarget = cols.reduce((a, c) => a + c.w, 0);
  const scale = pageW / totalTarget;
  let scaled = cols.map((c) => ({ ...c, w: Math.floor(c.w * scale) }));
  const sum = scaled.reduce((a, c) => a + c.w, 0);
  scaled[scaled.length - 1].w += pageW - sum; // exact fit

  return scaled;
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

      const newPage = () => {
        doc.addPage();
        pageNo += 1;
        drawChrome(doc, title, logo, pageNo);
        doc.y = 60;
      };

      // Page 1 chrome
      drawChrome(doc, title, logo, pageNo);
      doc.y = 60;

      const L = doc.page.margins.left;
      const pageW = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      const bottomLimit = () => doc.page.height - doc.page.margins.bottom - 60;

      // --- TABLE FIRST (always on page 1) ---
      doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Data");
      doc.moveDown(0.5);

      if (!rows.length) {
        doc.font("Helvetica").fontSize(10).fillColor(BRAND.subText).text(
          "No rows were provided to export. (Your export request must include body.rows as an array.)",
          { width: pageW }
        );
      } else {
        const cols = buildColumnsFromRows(rows, pageW);
        const rowH = 18;

        // x positions for perfect alignment
        const xPos = [L];
        for (let i = 0; i < cols.length; i++) xPos.push(xPos[i] + cols[i].w);

        const drawCellText = (text, x, y, w, isHeader) => {
          const padX = 4;
          const padY = 5;
          doc.save();
          doc.rect(x + 1, y + 1, w - 2, rowH - 2).clip();
          doc.font(isHeader ? "Helvetica-Bold" : "Helvetica").fontSize(9).fillColor("#111827");
          doc.text(safe(text), x + padX, y + padY, {
            width: w - padX * 2,
            lineBreak: false,
            ellipsis: true,
          });
          doc.restore();
        };

        const drawGrid = (y) => {
          doc.save();
          doc.strokeColor(BRAND.line).strokeOpacity(0.45);
          // verticals
          for (let i = 0; i < xPos.length; i++) {
            doc.moveTo(xPos[i], y).lineTo(xPos[i], y + rowH).stroke();
          }
          // bottom
          doc.moveTo(L, y + rowH).lineTo(L + pageW, y + rowH).stroke();
          doc.restore();
        };

        const drawHeader = () => {
          const y = doc.y;

          doc.save();
          doc.fillColor(BRAND.tableHeaderFill);
          doc.rect(L, y, pageW, rowH).fill();
          doc.restore();

          for (let i = 0; i < cols.length; i++) {
            drawCellText(cols[i].label, xPos[i], y, cols[i].w, true);
          }
          drawGrid(y);

          doc.y = y + rowH;
        };

        const drawRow = (r, idx) => {
          const y = doc.y;

          if (idx % 2 === 1) {
            doc.save();
            doc.fillColor(BRAND.zebraFill);
            doc.rect(L, y, pageW, rowH).fill();
            doc.restore();
          }

          for (let i = 0; i < cols.length; i++) {
            const c = cols[i];
            drawCellText(r?.[c.key] ?? "", xPos[i], y, c.w, false);
          }
          drawGrid(y);

          doc.y = y + rowH;
        };

        // Ensure table starts right under header
        drawHeader();

        const MAX = 1500;
        for (let i = 0; i < Math.min(rows.length, MAX); i++) {
          if (doc.y > bottomLimit() - rowH) {
            newPage();
            doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Data (continued)");
            doc.moveDown(0.5);
            drawHeader();
          }
          drawRow(rows[i], i);
        }

        if (rows.length > MAX) {
          if (doc.y > bottomLimit() - 30) newPage();
          doc.moveDown(1);
          doc.font("Helvetica").fontSize(9).fillColor(BRAND.subText).text(`Showing first ${MAX} of ${rows.length} rows.`);
        }
      }

      // --- ANSWER AFTER (page 2) ---
      newPage();

      // optional meta on answer page
      if (meta) {
        doc.font("Helvetica-Bold").fontSize(11).fillColor("#111827").text("Parameters");
        doc.moveDown(0.3);
        doc.font("Helvetica").fontSize(10).fillColor("#111827");
        for (const k of Object.keys(meta)) {
          doc.fillColor(BRAND.subText).font("Helvetica-Bold").text(`${k}: `, { continued: true });
          doc.fillColor("#111827").font("Helvetica").text(safe(meta[k]));
          if (doc.y > bottomLimit()) newPage();
        }
        doc.moveDown(0.8);
      }

      doc.font("Helvetica-Bold").fontSize(12).fillColor("#111827").text("Answer");
      doc.moveDown(0.5);
      doc.font("Helvetica").fontSize(10).fillColor("#111827");
      doc.text(answer, { width: pageW });

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

