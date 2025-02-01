const express = require("express");
const ExcelJS = require("exceljs");
const router = express.Router();

router.post("/export", async (req, res) => {
  try {
    console.log("✅ Export API Hit!"); // Confirm route is hit
    // console.log("🔍 Received Data:", req.body); // Log request body

    const { sheetData, sheetNames } = req.body;

    if (!sheetData || !sheetNames) {
      console.error("❌ Missing data in request body");
      return res.status(400).json({ error: "No data provided for export." });
    }

    const workbook = new ExcelJS.Workbook();

    for (const sheetName of sheetNames) {
      const worksheet = workbook.addWorksheet(sheetName);
      if (sheetData[sheetName] && sheetData[sheetName].length > 0) {
        // Get headers from the first row
        const headers = Object.keys(sheetData[sheetName][0]);
        worksheet.addRow(headers);
        // Add each row of data
        sheetData[sheetName].forEach(row => {
          const rowData = headers.map(header => row[header]);
          worksheet.addRow(rowData);
        });
      } else {
        worksheet.addRow(["No data"]);
      }
    }
    res.setHeader("Content-Disposition", "attachment; filename=validated_data.xlsx");
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

    return workbook.xlsx.write(res).then(() => {
      res.end();
    });

  } catch (error) {
    console.error("❌ Error exporting data:", error);
    return res.status(500).json({ error: "Error exporting data." });
  }
});

module.exports = router;
