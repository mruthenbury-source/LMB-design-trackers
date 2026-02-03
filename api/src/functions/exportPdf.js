import { app } from "@azure/functions";
import PDFDocument from "pdfkit";

function safeStr(v) {
  return (v ?? "").toString();
}

function wrapLines(doc, text, maxWidth) {
  const out = [];
  const paras = safeStr(text).split("\n");
  for (const p of paras) {
    const wrapped = doc.splitTextToSize(p, maxWidth);
    out.push(...wrapped, "");
  }
  return out;
}

app.http("exportPdf", {
  methods: ["POST"],
  authLevel: "anonymous",
  handler: async (req) => {
    try {
      const body = await req.json().catch(() => ({}));

      const title = safeStr(body.title || "Chat Export");
      const answer = safeStr(body.answer || "");
      const rows = Array.isArray(body.rows) ? body.rows : [];
      const meta = body.meta && typeof body.meta === "object" ? body.meta : {};

      const doc = new PDFDocument({ size: "A4", margin: 36 });
      const chunks = [];
      doc.on("data", (c) => chunks.push(c));
      const done = new Promise((resolve) => doc.on("end", resolve));

      // Header
      doc.font("Helvetica-Bold").fontSize(16).text(title, { underline: true });
      doc.moveDown(0.4);
      doc.font("Helvetica").fontSize(9).fillColor("gray").text(`Generated: ${new Date().toISOString()}`);
      doc.fillColor("black");
      doc.moveDown(1);

      // Meta (optional)
      const metaKeys = Object.keys(meta);
      if (metaKeys.length) {
        doc.font("Helvetica-Bold").fontSize(11).text("Parameters", { underline: true });
        doc.moveDown(0.3);
        doc.font("Helvetica").fontSize(10);
        for (const k of metaKeys) {
          doc.text(`${k}: ${safeStr(meta[k])}`);
        }
        doc.moveDown(1);
      }

      // Answer
      doc.font("Helvetica-Bold").fontSize(11).text("Answer", { underline: true });
      doc.moveDown(0.3);
      doc.font("Helvetica").fontSize(10);
      const maxWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      const lines = wrapLines(doc, answer, maxWidth);

      for (const ln of lines) {
        if (doc.y > doc.page.height - doc.page.margins.bottom - 40) doc.addPage();
        doc.text(ln);
      }

      // Data table (optional)
      if (rows.length) {
        doc.addPage();
        doc.font("Helvetica-Bold").fontSize(11).text("Data", { underline: true });
        doc.moveDown(0.5);

        const cols = [
          { key: "project", label: "Project", w: 70 },
          { key: "supplier", label: "Supplier", w: 55 },
          { key: "responsibility", label: "Resp", w: 70 },
          { key: "item", label: "Item", w: 170 },
          { key: "requiredOnSite", label: "Req", w: 55 },
          { key: "statusA", label: "A", w: 55 },
        ];

        const startX = doc.page.margins.left;
        let y = doc.y;

        const drawRow = (vals, isHeader = false) => {
          let x = startX;
          doc.font(isHeader ? "Helvetica-Bold" : "Helvetica").fontSize(isHeader ? 8 : 8);
          for (let i = 0; i < vals.length; i++) {
            doc.text(safeStr(vals[i]), x, y, { width: cols[i].w, ellipsis: true });
            x += cols[i].w;
          }
          y += 12;
          if (y > doc.page.height - doc.page.margins.bottom - 20) {
            doc.addPage();
            y = doc.page.margins.top;
          }
        };

        drawRow(cols.map((c) => c.label), true);
        doc.moveTo(startX, y).lineTo(startX + cols.reduce((a, c) => a + c.w, 0), y).stroke();
        y += 6;

        const limit = 1000;
        for (const r of rows.slice(0, limit)) {
          drawRow(cols.map((c) => r?.[c.key] ?? ""));
        }

        if (rows.length > limit) {
          doc.font("Helvetica").fontSize(9).text(`Showing first ${limit} of ${rows.length} rows.`, startX, y);
        }
      }

      doc.end();
      await done;

      const pdfBuffer = Buffer.concat(chunks);

      return {
        status: 200,
        headers: {
          "Content-Type": "application/pdf",
          "Content-Disposition": 'attachment; filename="chat-export.pdf"',
        },
        body: pdfBuffer,
      };
    } catch (e) {
      return { status: 500, jsonBody: { error: "exportPdf failed", details: String(e) } };
    }
  },
});
