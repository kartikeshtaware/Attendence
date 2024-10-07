const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json()); // To parse JSON bodies

// Set up multer storage configuration with a fixed filename (e.g., attendance.xlsx)
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/'); // Set the destination folder
  },
  filename: (req, file, cb) => {
    cb(null, 'attendance.xlsx'); // Always save as 'attendance.xlsx'
  }
});

// Set up multer with the storage configuration
const upload = multer({ storage });

// Serve static files from the 'public' directory
app.use(express.static('public'));

// Global variable to store the file path of the uploaded file
const uploadedFilePath = path.join(__dirname, 'uploads', 'attendance.xlsx');

// Function to convert row and column index to Excel-style cell address
function getExcelCellAddress(row, column) {
  const letters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let columnName = '';

  while (column > 0) {
    columnName = letters[(column - 1) % 26] + columnName;
    column = Math.floor((column - 1) / 26);
  }

  return columnName + row;
}

// Route to upload Excel file
app.post('/upload-excel', upload.single('file'), (req, res) => {
  try {
    // Check if file was uploaded
    if (!req.file) {
      return res.status(400).json({ message: 'No file uploaded.' });
    }

    // Reading the uploaded Excel file
    const workbook = XLSX.readFile(uploadedFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    res.json({
      message: 'File uploaded successfully',
      sheetData, // Send the sheet data to the frontend
      filePath: uploadedFilePath // Include the file path for later reference
    });
  } catch (error) {
    res.status(500).json({ message: 'Error reading Excel file.', error });
  }
});

// Route to update the same original Excel file based on QR code scan and date
app.post('/update-excel', (req, res) => {
  const { qrSubstring, attendanceDate } = req.body; // Expecting both date and substring

  // Check if uploaded file path is valid
  if (!uploadedFilePath || !fs.existsSync(uploadedFilePath)) {
    return res.status(400).json({ message: 'No Excel file uploaded.' });
  }

  try {
    // Read the uploaded file again from its original path
    const workbook = XLSX.readFile(uploadedFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Convert sheet to array of arrays for easier processing
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Find the index of the date in the first row (header)
    const dateRow = sheetData[0]; // First row with headers
    const dateColumnIndex = dateRow.findIndex(date => date === attendanceDate); // Find the index of the specified date

    if (dateColumnIndex === -1) {
      return res.status(404).json({ message: 'Date not found in the Excel sheet.' });
    }

    let matched = false;

    // Find the row where the first column matches the QR substring (matching the name)
    for (let i = 1; i < sheetData.length; i++) {
      const excelValue = sheetData[i][0]?.toString().trim().toLowerCase(); // Normalize Excel data
      const qrValue = qrSubstring.toLowerCase(); // Normalize QR data for comparison

      if (excelValue && excelValue.startsWith(qrValue)) {
        // Assuming the cell for attendance is in the found date column
        const rowIndex = i + 1; // Row in Excel (1-based index)
        const columnIndex = dateColumnIndex + 1; // Adjust for 1-based index in Excel

        // Update the value with 'P' in the found date column
        const cellAddress = getExcelCellAddress(rowIndex, columnIndex);
        worksheet[cellAddress] = { v: 'P' }; // Mark attendance with 'P'

        matched = true;
        break;
      }
    }

    if (!matched) {
      return res.status(404).json({ message: 'No match found in the Excel sheet for the name.' });
    }

    // Save the updated Excel file (overwrite the original file)
    XLSX.writeFile(workbook, uploadedFilePath);

    res.json({
      message: `Updated attendance for ${qrSubstring} on ${attendanceDate}`,
      downloadLink: `/uploads/attendance.xlsx` // Return link for updated file
    });
  } catch (error) {
    res.status(500).json({ message: 'Error updating Excel file.', error });
  }
});

// Start the server
app.listen(3000, () => {
  console.log('Server running on http://localhost:3000');
});
