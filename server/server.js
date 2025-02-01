// server.js
const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const mongoose = require('mongoose');
const cors = require('cors');
const sheetConfig = require('./SheetConfig');
const Record = require('./models/Record');
const exportRoutes = require('./routes/export');
const dotenv = require('dotenv');

dotenv.config();
const mongoURI = process.env.MONGODB_URI;

const app = express();
const port = process.env.PORT || 5000;

// Middleware
app.use(express.json());
app.use(cors());

// Multer configuration to store file in memory
const storage = multer.memoryStorage();
const upload = multer({
  storage: storage,
  limits: { fileSize: 2 * 1024 * 1024 } // 2 MB limit
});

// Connect to MongoDB Atlas
mongoose
  .connect(mongoURI, {
    useNewUrlParser: true,
    useUnifiedTopology: true
  })
  .then(() => console.log('Connected to MongoDB Atlas'))
  .catch((err) => console.error('MongoDB connection error:', err));

/**
 * Helper: Validate a single row based on the provided config.
 * @param {Object} rowData - Object with keys as Excel column headers.
 * @param {Object} config - The sheet config from sheetConfig.js.
 * @param {Number} rowNumber - The row number in the Excel sheet.
 * @returns {Array} errors - Array of error messages (if any).
 */
function validateRow(rowData, config, rowNumber) {
  const errors = [];
  const rules = config.validationRules;
  const now = new Date();
  const currentMonth = now.getMonth();
  const currentYear = now.getFullYear();

  for (const col in rules) {
    const rule = rules[col];
    const value = rowData[col];

    // Check required fields
    if (rule.required && (value === undefined || value === null || value === '')) {
      errors.push(`Row ${rowNumber}: "${col}" is required.`);
      continue;
    }
    if (value === undefined || value === null || value === '') {
      // If not required and empty, skip further validations for this field.
      continue;
    }
    // Validate based on type
    if (rule.type === 'number') {
      const num = parseFloat(value);
      if (isNaN(num)) {
        errors.push(`Row ${rowNumber}: "${col}" must be numeric.`);
      } else if (rule.min !== undefined && num < rule.min) {
        errors.push(`Row ${rowNumber}: "${col}" must be greater than ${rule.min}.`);
      }
    } else if (rule.type === 'date') {
      const date = new Date(value);
      if (isNaN(date.getTime())) {
        errors.push(`Row ${rowNumber}: "${col}" must be a valid date.`);
      } else if (rule.currentMonth) {
        if (date.getMonth() !== currentMonth || date.getFullYear() !== currentYear) {
          errors.push(`Row ${rowNumber}: "${col}" must be within the current month.`);
        }
      }
    } else if (rule.type === 'string') {
      // For strings, you might add more validations if needed.
      if (typeof value !== 'string') {
        errors.push(`Row ${rowNumber}: "${col}" must be a string.`);
      }
    }
    // Validate allowed values if provided (like for Verified)
    if (rule.allowedValues && !rule.allowedValues.includes(value)) {
      errors.push(`Row ${rowNumber}: "${col}" must be one of ${rule.allowedValues.join(', ')}.`);
    }
  }
  return errors;
}

app.use('/api', exportRoutes);

/**
 * Endpoint to process the uploaded .xlsx file.
 */
app.post('/api/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded.' });
    }

    // Use ExcelJS to process the file
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(req.file.buffer);

    const sheetNames = [];
    const sheetData = {};
    const validationErrors = {};

    // Process each worksheet
    workbook.eachSheet((worksheet, sheetId) => {
      const sheetName = worksheet.name;
      sheetNames.push(sheetName);

      // Use the default config for this example.
      // In the future you could pick a config based on the sheet name.
      const config = sheetConfig.default;
      const headers = {};

      // Read header row (assumed to be the first row)
      const headerRow = worksheet.getRow(1);
      headerRow.eachCell((cell, colNumber) => {
        headers[colNumber] = cell.value;
      });

      // Check if required columns exist
      const missingColumns = [];
      Object.keys(config.columnMapping).forEach((expectedColumn) => {
        if (!Object.values(headers).includes(expectedColumn)) {
          missingColumns.push(expectedColumn);
        }
      });
      if (missingColumns.length > 0) {
        validationErrors[sheetName] = [
          {
            row: 1,
            error: `Missing required columns: ${missingColumns.join(', ')}`
          }
        ];
        // Skip further processing for this sheet if headers are missing.
        return;
      }

      // Process data rows
      sheetData[sheetName] = [];
      validationErrors[sheetName] = [];
      worksheet.eachRow((row, rowNumber) => {
        // Skip header row
        if (rowNumber === 1) return;

        // Build a row object using header names as keys.
        const rowObject = {};
        row.eachCell((cell, colNumber) => {
          const header = headers[colNumber];
          rowObject[header] = cell.value;
        });

        // Validate the row
        const rowErrors = validateRow(rowObject, config, rowNumber);
        if (rowErrors.length > 0) {
          rowErrors.forEach((errorMsg) => {
            validationErrors[sheetName].push({ row: rowNumber, error: errorMsg });
          });
        }

        // Map the row using the column mapping to database fields.
        const mappedRow = {};
        Object.entries(config.columnMapping).forEach(([excelColumn, dbField]) => {
          mappedRow[dbField] = rowObject[excelColumn];
          // For date fields, convert to a Date object if valid.
          if (config.validationRules[excelColumn].type === 'date' && rowObject[excelColumn]) {
            mappedRow[dbField] = new Date(rowObject[excelColumn]);
          }
          // For numeric fields, convert to a number.
          if (config.validationRules[excelColumn].type === 'number' && rowObject[excelColumn]) {
            mappedRow[dbField] = parseFloat(rowObject[excelColumn]);
          }
        });
        sheetData[sheetName].push(mappedRow);
      });

      // Remove validation error array if no errors were found for this sheet.
      if (validationErrors[sheetName].length === 0) {
        delete validationErrors[sheetName];
      }
    });

    // Return the parsed data, sheet names, and any validation errors.
    return res.json({
      sheetNames,
      sheetData,
      validationErrors
    });
  } catch (err) {
    console.error('Error processing file:', err);
    return res.status(500).json({ error: 'Error processing file.' });
  }
});

/**
 * Endpoint to import valid rows into MongoDB.
 * Expects a JSON body with "data": { sheetName: [ ... rows ... ] }
 * This endpoint will import only valid rows (based on previously validated data).
 */
// In your server.js (or wherever your import endpoint is defined)
app.post("/api/import", async (req, res) => {
  try {
    // Accept data from either 'data' or 'sheetData'
    const data = req.body.data || req.body.sheetData;
    // Accept errors (default to empty object if not provided)
    const errors = req.body.errors || {};

    if (!data) {
      return res.status(400).json({ error: "No data provided for import." });
    }

    let importedCount = 0;
    // Loop over each sheet in the data
    for (const sheet in data) {
      // If the sheet data is undefined or null, skip it.
      if (!data[sheet]) continue;

      // Build a set of row numbers that have errors (if any)
      const errorRows = new Set((errors[sheet] || []).map((err) => err.row));

      // Filter out rows that either have errors or are missing the required 'name' field.
      const validRows = data[sheet].filter((row, index) => {
        // Assuming the first row is the header, data starts at row 2.
        const rowNumber = index + 2;
        return !errorRows.has(rowNumber) && row.name;
      });

      if (validRows.length > 0) {
        // Insert valid rows into MongoDB. (Record is your Mongoose model.)
        const result = await Record.insertMany(validRows);
        importedCount += result.length;
      }
    }

    res.json({
      message: `Successfully imported ${importedCount} records.`,
    });
  } catch (err) {
    console.error("Error importing data:", err);
    res.status(500).json({ error: "Error importing data." });
  }
});

const path = require("path");
// Serve static files from the React app
app.use(express.static(path.join(__dirname, "../client/build")));

// The "catchall" handler: for any request that doesn't match an API route,
// send back React's index.html file.
app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "../client/build", "index.html"));
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});