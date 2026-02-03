import { app } from "@azure/functions";
import PDFDocument from "pdfkit";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function safe(v) {
  if (v === null || v === undefined) return "";
  return typeof v === "string" ? v : JSON.stringify(v);
}

function loadLogo() {
  try {
    return fs.readFileSync(
      path.join(__dirname, "..", "assets", "logo.png")
    );
  } catch {
    return null;
  }
}

function headerFooter(doc, title, logo, pageNo) {
  const L = doc.page.margins.left;
  const R = doc.page.width - doc.page.margins.right;
  const T = doc.page.margins.top;
  const B = doc.page.height - doc.page.margins.bottom;

  // Header line
  doc.moveTo(L, T - 10).lineTo(R, T - 10).stroke();

  if (logo) doc.image(logo, L, T - 36, { width: 80 });

  doc.font("Helvetica-Bold")
    .fontSize(14)
    .text(title, L + (logo ? 95 : 0), T - 34);

  // Footer
  doc.moveTo(L, B + 10).lineTo(R, B + 10).stroke();

  doc.fontSize(9).fillColor("gray");
  doc.text(`Generated ${new Date().toLocaleString()}`, L, B + 16);
  doc.text(`Page ${pageNo}`, L, B + 16, { width: R - L, align: "right" });
  doc.fillColor("black");
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

      const doc = new PDFDocument({ margin: 40 });
      const buffers = [];
      doc.on("data", (d) => buffers.push(d));
      const done = new Promise((r) => doc.on("end", r));

      const logo = loadLogo();
      let pageNo = 1;

      headerFooter(doc, title, logo, pageNo);
      doc.on("pageAdded", () => {
        pageNo++;
        headerFooter(doc, title, logo, pageNo);
      });

      // ===== ANSWER TEXT =====
      doc.moveDown(2);
      doc.fontSize(12).font("Helvetica-Bold").text("Answer");
      doc.moveDown(0.5);
      doc.fontSize(10).font("Helvetica").text(answer, { width: 500 });

      // ===== TABLE =====
      if (rows.length) {
        doc.addPage();

        const pageW =
          doc.page.width - doc.page.margins.left - doc.page.margins.right;

        const cols = [
          { key: "project", label: "Project", w: 80 },
          { key: "supplier", label: "Supplier", w: 70 },
          { key: "responsibility", label: "Resp", w: 70 },
          {
            key: "item",
            label: "Item",
            w: pageW - (80 + 70 + 70 + 55 + 55),
          },
          { key: "requiredOnSite", label: "Req", w: 55 },
          { key: "statusA", label: "A", w: 55 },
        ];

        const startX = doc.page.margins.left;
        let y = doc.y + 10;
        const rowH = 14;
        const totalW = cols.reduce((a, c) => a + c.w, 0);

        const drawRow = (vals, header = false) => {
          let x = startX;
          doc.font(header ? "Helvetica-Bold" : "Helvetica").fontSize(8);

          vals.forEach((v, i) => {
            doc.text(safe(v), x, y, {
              width: cols[i].w,
              ellipsis: true,
            });
            x += cols[i].w;
          });

          // Row line
          doc
            .moveTo(startX, y + rowH - 2)
            .lineTo(startX + totalW, y + rowH - 2)
            .strokeOpacity(0.2)
            .stroke()
            .strokeOpacity(1);

          y += rowH;

          if (y > doc.page.height - doc.page.margins.bottom - 40) {
            doc.addPage();
            y = doc.page.margins.top + 20;
          }
        };

        // Header row
        drawRow(cols.map((c) => c.label), true);

        rows.slice(0, 1000).forEach((r) =>
          drawRow(cols.map((c) => r?.[c.key] ?? ""))
        );
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


