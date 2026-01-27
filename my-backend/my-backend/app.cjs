const express = require("express");
const multer = require("multer");
const path = require("path");
const fs = require("fs").promises; // Use promises version
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();
const PORT = 3000;

// === Config ===
const UPLOAD_DIR = path.join(__dirname, "uploads");
const FILE_PATH = path.join(UPLOAD_DIR, "schedule.xlsx");

// Ensure upload directory exists
await fs.mkdir(UPLOAD_DIR, { recursive: true }).catch(console.error);

app.use(cors({
  origin: "http://localhost:5500", // or "*" if you want anyone
  methods: ["GET", "POST"],
}));
app.use(express.json());

// === Multer: only accept Excel files ===
const upload = multer({
  storage: multer.diskStorage({
    destination: (req, file, cb) => cb(null, UPLOAD_DIR),
    filename: (req, file, cb) => cb(null, "schedule.xlsx"),
  }),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10 MB max
  fileFilter: (req, file, cb) => {
    const allowed = [
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // .xlsx
      "application/vnd.ms-excel" // .xls
    ];
    if (allowed.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error("Only .xlsx and .xls files are allowed!"));
    }
  }
});

// === In-memory cache ===
let scheduleData = [];
let lastLoadTime = null;

// Load and parse Excel file
async function loadSchedule() {
  try {
    if (!(await fs.stat(FILE_PATH).catch(() => false))) {
      console.log("No schedule file found yet.");
      scheduleData = [];
      return;
    }

    const workbook = XLSX.readFile(FILE_PATH);
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

    if (rows.length <= 1) {
      scheduleData = [];
      console.log("Excel file is empty or has only headers.");
      return;
    }

    const dataRows = rows.slice(1);
    scheduleData = dataRows.map(r => ({
      Name: String(r[0] || "").trim(),
      Day: String(r[1] || "").trim(),
      Time: String(r[2] || "").trim(),
      Activity: String(r[3] || "").trim(),
      Description: String(r[4] || "").trim(),
    }));

    lastLoadTime = new Date();
    console.log(`Schedule loaded: ${scheduleData.length} entries at ${lastLoadTime.toLocaleTimeString()}`);
  } catch (err) {
    console.error("Failed to load schedule:", err.message);
    scheduleData = [];
  }
}

// === Routes ===

// Upload new schedule
app.post("/admin/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No file uploaded." });
  }

  await loadSchedule(); // Reload immediately

  res.json({
    message: "Schedule uploaded and updated successfully!",
    entries: scheduleData.length,
    uploadedAt: new Date().toLocaleString()
  });
}, (error, req, res, next) => {
  // Multer error handler
  res.status(400).json({ error: error.message });
});

// Get schedule by name
app.get("/schedule", async (req, res) => {
  const name = (req.query.name || "").toString().trim();

  if (!name) {
    return res.status(400).send("Missing or empty ?name= parameter");
  }

  // Reload if file changed (optional hot-reload)
  try {
    const stats = await fs.stat(FILE_PATH);
    if (lastLoadTime && stats.mtime > lastLoadTime) {
      console.log("File changed on disk → reloading...");
      await loadSchedule();
    }
  } catch (_) { /* file missing is okay */ }

  if (scheduleData.length === 0) {
    await loadSchedule();
    if (scheduleData.length === 0) {
      return res.status(404).send("No schedule data available yet. Please upload an Excel file first.");
    }
  }

  const lowerName = name.toLowerCase();
  const matches = scheduleData.filter(row => row.Name.toLowerCase() === lowerName);

  if (matches.length === 0) {
    return res.status(404).send(`No schedule found for "${name}".`);
  }

  const escape = (str) => String(str).replace(/[&<>"']/g, 
    c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[c]);

  const rows = matches.map(r => `
    <tr>
      <td>${escape(r.Day)}</td>
      <td>${escape(r.Time)}</td>
      <td>${escape(r.Activity)}</td>
      <td>${escape(r.Description)}</td>
    </tr>
  `).join("");

  const html = `
    <style>
      body { font-family: system-ui, sans-serif; padding: 1rem; }
      table { width: 100%; border-collapse: collapse; margin-top: 1rem; }
      th, td { border: 1px solid #666; padding: 0.8rem; text-align: left; }
      th { background: #f0f0f0; }
      tr:nth-child(even) { background: #fafafa; }
    </style>
    <h2>Schedule for ${escape(matches[0].Name)}</h2>
    <p><strong>${matches.length}</strong> entr${matches.length > 1 ? "ies" : "y"} found.</p>
    <table>
      <thead><tr>
        <th>Day</th>
        <th>Time</th>
        <th>Activity</th>
        <th>Description</th>
      </tr></thead>
      <tbody>${rows}</tbody>
    </table>
  `;

  res.send(html);
});