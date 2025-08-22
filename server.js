// server.js
const express = require("express");
const cors = require("cors");
const helmet = require("helmet");
const rateLimit = require("express-rate-limit");
const multer = require("multer");
const fs = require("fs");
const path = require("path");

const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const mammoth = require("mammoth");

const app = express();

// --- Security & basics ---
app.use(helmet());
app.use(cors({ origin: true }));       // tighten to your domain in production
app.use(express.json({ limit: "1mb" })); // small limit is fine for text
app.use("/",
  rateLimit({ windowMs: 60 * 1000, max: 60 }) // 60 req/min per IP
);

// serve a simple demo page
app.use(express.static(path.join(__dirname, "public")));

// In-memory uploads (keeps it lightweight)
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 5 * 1024 * 1024 }, // 5MB
  fileFilter: (req, file, cb) => {
    const okType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    const extOk = file.originalname.toLowerCase().endsWith(".docx");
    if (file.mimetype === okType && extOk) return cb(null, true);
    cb(new Error("Only .docx files are allowed"));
  },
});

// Helper to render a docx from a template buffer + data
function renderDocx(templateBuffer, data) {
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.setData(data);
  try {
    doc.render();
  } catch (e) {
    // Helpful error to debug missing/typo placeholders
    const msg = e.properties?.errors?.map(err => err.properties?.explanation).join("\n") || e.message;
    throw new Error("Template render error:\n" + msg);
  }
  return doc.getZip().generate({ type: "nodebuffer", compression: "DEFLATE" });
}

// --- 1) FORM MODE: generate from your default template ---
app.post("/api/generate", async (req, res) => {
  try {
    // Example expected fields; add/remove as you like
    const {
      refNo = "",
      date = "",
      recipientName = "",
      recipientAddress = "",
      subject = "",
      content = "",
      senderName = "",
      senderPosition = "",
      organization = "",
    } = req.body || {};

    // Load your default template
    const templatePath = path.join(__dirname, "templates", "default-letter.docx");
    const templateBuffer = fs.readFileSync(templatePath);

    const buffer = renderDocx(templateBuffer, {
      refNo,
      date,
      recipientName,
      recipientAddress,
      subject,
      content,
      senderName,
      senderPosition,
      organization,
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="Official_Letter.docx"');
    return res.status(200).send(buffer);
  } catch (err) {
    return res.status(400).json({ error: err.message });
  }
});

// --- 2) UPLOAD MODE: user uploads a .docx template, we fill and return ---
app.post("/api/upload-and-generate", upload.single("template"), async (req, res) => {
  try {
    if (!req.file) throw new Error("No .docx template uploaded");
    // The form should send a JSON string named "data"
    const data = req.body?.data ? JSON.parse(req.body.data) : {};

    const buffer = renderDocx(req.file.buffer, data);

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    );
    res.setHeader("Content-Disposition", 'attachment; filename="Generated_From_Template.docx"');
    return res.status(200).send(buffer);
  } catch (err) {
    return res.status(400).json({ error: err.message });
  }
});

// --- 3) INSPECT: discover placeholders in an uploaded template (optional UX helper) ---
app.post("/api/inspect-template", upload.single("template"), async (req, res) => {
  try {
    if (!req.file) throw new Error("No .docx template uploaded");
    const { value } = await mammoth.extractRawText({ buffer: req.file.buffer });
    const tags = Array.from(value.matchAll(/\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}/g)).map(m => m[1]);
    const unique = [...new Set(tags)];
    return res.json({ placeholders: unique });
  } catch (err) {
    return res.status(400).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running http://localhost:${PORT}`));
