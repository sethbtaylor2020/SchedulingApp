// server.js
// Personal schedule app
// - Loads schedule.xlsx from ./uploads on startup
// - Lookup: http://localhost:3000/?name=John
// - Always shows reference.pdf in an iframe viewer (no download prompt)

const express = require("express");
const path = require("path");
const fs = require("fs");
const XLSX = require("xlsx");
const cors = require("cors");

const app = express();
const PORT = 3000;

/* =========================
   MIDDLEWARE
========================= */
app.use(cors());

/* =========================
   PATHS
========================= */
const uploadDir = path.join(__dirname, "uploads");
const excelPath = path.join(uploadDir, "schedule.xlsx");
const pdfPath = path.join(uploadDir, "reference.pdf");

// Ensure uploads folder exists
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

let scheduleData = [];

/* =========================
   HELPERS
========================= */
function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function loadSchedule() {
  if (!fs.existsSync(excelPath)) {
    console.log("❌ schedule.xlsx not found at:", excelPath);
    scheduleData = [];
    return;
  }

  try {
    const workbook = XLSX.readFile(excelPath);
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    scheduleData = XLSX.utils.sheet_to_json(sheet);
    console.log(`✅ Loaded ${scheduleData.length} schedule rows from ${excelPath}`);
  } catch (err) {
    console.error("Failed to read schedule.xlsx:", err);
    scheduleData = [];
  }
}

// Load schedule on startup
loadSchedule();
// Serve the main lookup page (index.html)
app.get("/", (req, res) => {
  const indexPath = path.join(__dirname, "index.html");
  if (!fs.existsSync(indexPath)) {
    return res.status(500).send(`<h2>index.html not found</h2><p>Place your index.html in the project root.</p>`);
  }
  res.sendFile(indexPath);
});

// Serve reference.pdf inline
app.get("/reference.pdf", (req, res) => {
  if (!fs.existsSync(pdfPath)) {
    return res.status(404).send("reference.pdf not found in ./uploads/");
  }
  res.setHeader("Content-Type", "application/pdf");
  res.setHeader("Content-Disposition", 'inline; filename="reference.pdf"');
  res.sendFile(pdfPath);
});

// PDF viewer page (iframe target)
app.get("/pdfviewer", (req, res) => {
  res.send(`
<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>Reference PDF</title>
  <style>
    html, body { margin:0; padding:0; height:100%; width:100%; overflow:hidden; background:#333; }
    embed { width:100%; height:100%; border:none; }
  </style>
</head>
<body>
  <embed src="/reference.pdf#view=FitH&toolbar=1&navpanes=0&scrollbar=1" type="application/pdf">
</body>
</html>
  `);
});
/* =========================
   ROUTES
========================= */

// Schedule lookup - shows ALL rows for the person, grouped by day
app.get("/schedule", (req, res) => {
  const nameQuery = String(req.query.name || "").trim().toLowerCase();

  let matchingRows = scheduleData;

  if (nameQuery) {
    matchingRows = scheduleData.filter(row => 
      String(row.Name || "").toLowerCase().includes(nameQuery)
    );

    if (matchingRows.length === 0) {
      return res.status(404).type("html")
        .send(`<span class="error">No schedule found for <strong>${escapeHtml(req.query.name)}</strong>.</span>`);
    }
  }

  // Group by Day
  const days = {};
  matchingRows.forEach(row => {
    const day = row.Day || "Unknown";
    if (!days[day]) days[day] = [];
    days[day].push(row);
  });

  // Sort days
  const dayOrder = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];
  const sortedDays = Object.keys(days).sort((a, b) => 
    dayOrder.indexOf(a) - dayOrder.indexOf(b)
  );

  let html = `<div style="font-family: system-ui, sans-serif;">`;

  if (nameQuery) {
    html += `<h2>Schedule for <strong>${escapeHtml(req.query.name)}</strong> (${matchingRows.length} shift${matchingRows.length > 1 ? 's' : ''})</h2>`;
  } else {
    html += `<h2>Full Schedule</h2>`;
  }

  sortedDays.forEach(day => {
    html += `<h3 style="color: #0066ff; margin: 32px 0 8px;">${escapeHtml(day)}</h3>`;
    html += `<table border="1" cellpadding="10" cellspacing="0" style="border-collapse: collapse; width: 100%; background: #fff;">
      <tr><th>Time</th><th>Activity</th><th>Description</th><th>Location</th></tr>`;

    days[day].forEach(row => {
      html += `<tr>
        <td><strong>${escapeHtml(row.Time || "")}</strong></td>
        <td>${escapeHtml(row.Activity || "")}</td>
        <td>${escapeHtml(row.Description || "")}</td>
        <td>${escapeHtml(row.Location || "")}</td>
      </tr>`;
    });

    html += `</table>`;
  });

  html += `
  <hr style="margin: 32px 0;" />
  <h3>Event Details</h3>
  <iframe src="/pdfviewer" style="width: 100%; height: 800px; border: 1px solid #ddd; border-radius: 8px;" loading="lazy"></iframe>
  </div>`;

  res.type("html").send(html);
});

/* =========================
   START SERVER
========================= */
app.listen(PORT, () => {
  console.log(`🚀 Server running at http://localhost:${PORT}`);
  console.log(`   Update schedule: replace ./uploads/schedule.xlsx and restart server`);
});