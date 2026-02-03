import { app } from "@azure/functions";
import PDFDocument from "pdfkit";

function safeStr(v) {
  if (v === null || v === undefined) return "";
  return typeof v === "string" ? v : JSON.stringify(v, null, 2);
}

function wrapTextToWidth(doc, text, maxWidth) {
  // Simple deterministic wrapper using pdfkit widthOfString
  const out = [];
  const paragraphs = safeStr(text).split("\n");

  for (const paragraph of paragraphs) {
    const words = paragraph.split(/\s+/).filter(Boolean);
    if (words.length === 0) {
      out.push(""); // blank line
      continue;
    }

    let line = "";
    for (const w of words) {
      const candidate = line ? `${line} ${w}` : w;
      const width = doc.widthOfString(candidate);

      if (width <= maxWidth) {
        line = candidate;
      } else {
        if (line) out.push(line);
        // Handle very long single words (no spaces)
        if (doc.widthOfString(w) > maxWidth) {
          let chunk = "";
          for (const ch of w) {
            const cand2 = chunk + ch;
            if (doc.widthOfString(cand2) <= maxWidth) chunk = cand2;
            else {
              if (chunk) out.push(chunk);
              chunk = ch;
            }
          }
          line = chunk;
        } else {
          line = w;
        }
      }
    }
    if (line) out.push(line);
    out.push(""); // paragraph spacing
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
      const meta = body.meta || {};

      const doc = new PDFDocument({ size: "A4", margin: 36 });
      const chunks = [];
      doc.on("data", (c) => chunks.push(c));
      const done = new Promise((resolve) => doc.on("end", resolve));

      // Header
      doc.fontSize(16).font("Helvetica-Bold").text(title, { underline: true });
      doc.moveDown(0.3);
      doc.fontSize(9).font("Helvetica").fillColor("gray").text(`Generated: ${new Date().toISOString()}`);
      doc.fillColor("black");
      doc.moveDown(1);

      // Meta
      const metaKeys = Object.keys(meta || {});
      if (metaKeys.length) {
        doc.fontSize(11).font("Helvetica-Bold").text("Parameters", { underline: true });
        doc.moveDown(0.2);
        doc.fontSize(10).font("Helvetica");
        for (const k of metaKeys) {
          doc.text(`${k}: ${safeStr(meta[k])}`);
        }
        doc.moveDown(0.8);
      }

      // Answer
      doc.fontSize(11).font("Helvetica-Bold").text("Answer", { underline: true });
      doc.moveDown(0.2);
      doc.fontSize(10).font("Helvetica");

      const maxWidth = doc.page.width - doc.page.margins.left - doc.page.margins.right;
      const lines = wrapTextToWidth(doc, answer, maxWidth);

      for (const ln of lines) {
        if (doc.y > doc.page.height - doc.page.margins.bottom - 40) doc.addPage();
        doc.text(ln);
      }

      // Rows table (optional)
      if (rows.length) {
        doc.addPage();
        doc.fontSize(11).font("Helvetica-Bold").text("Data", { underline: true });
        doc.moveDown(0.5);
        doc.fontSize(8).font("Helvetica");

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

        const rowWidth = cols.reduce((a, c) => a + c.w, 0);

        const drawRow = (vals, isHeader = false) => {
          let x = startX;
          doc.font(isHeader ? "Helvetica-Bold" : "Helvetica");
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
        doc.moveTo(startX, y).lineTo(startX + rowWidth, y).stroke();
        y += 6;

        const MAX = 1000;
        rows.slice(0, MAX).forEach((r) => drawRow(cols.map((c) => r?.[c.key] ?? "")));

        if (rows.length > MAX) {
          doc.font("Helvetica").text(`Showing first ${MAX} of ${rows.length} rows.`, startX, y);
        }
      }

      doc.end();
      await done;

      const pdfBuffer = Buffer.concat(chunks);

      return {
        status: 200,
        headers: {
          "Content-Type": "application/pdf",
          "Content-Disposition": `attachment; filename="chat-export.pdf"`,
        },
        body: pdfBuffer,
      };
    } catch (e) {
      return { status: 500, jsonBody: { error: "exportPdf failed", details: String(e) } };
    }
  },
});

