const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(express.json());

// Set up multer storage configuration
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, 'attendance.xlsx');
  }
});

// Set up multer with the storage configuration
const upload = multer({ storage });

// Serve static files from the 'public' directory
app.use(express.static('public'));

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

// Function to format date as YYYY-MM-DD
function formatDate(date) {
  const year = date.getFullYear();
  const month = (date.getMonth() + 1).toString().padStart(2, '0');
  const day = date.getDate().toString().padStart(2, '0');
  return `${year}-${month}-${day}`;
}

// Route to upload Excel file
app.post('/upload-excel', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ message: 'No file uploaded.' });
    }

    const filePath = path.join(__dirname, 'uploads', req.file.originalname);
    // res.json({ message: 'File uploaded successfully', filePath });
  } catch (error) {
    res.status(500).json({ message: 'Error uploading file.', error });
  }
});

// Route to update the Excel file for the selected subject
app.post('/update-excel', (req, res) => {
  const { qrSubstring, attendanceDate, subject } = req.body;
  const fileName = `${subject}.xlsx`; // Separate file for each subject

  const uploadedFilePath = path.join(__dirname, 'uploads', fileName);

  try {
    if (!fs.existsSync(uploadedFilePath)) {
      return res.status(400).json({ message: 'No Excel file for this subject uploaded.' });
    }

    const workbook = XLSX.readFile(uploadedFilePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const dateRow = sheetData[0];

    // Convert Excel date cells to YYYY-MM-DD format for comparison
    const formattedDateRow = dateRow.map(cell => {
      if (cell instanceof Date) {
        return formatDate(cell); // Format Date objects to 'YYYY-MM-DD'
      }
      return cell; // Return the value as-is if not a Date
    });

    const dateColumnIndex = formattedDateRow.findIndex(date => date === attendanceDate);

    if (dateColumnIndex === -1) {
      return res.status(404).json({ message: 'Date not found in the Excel sheet.' });
    }

    let matched = false;
    for (let i = 1; i < sheetData.length; i++) {
      const excelValue = sheetData[i][0]?.toString().trim().toLowerCase();
      const qrValue = qrSubstring.toLowerCase();

      if (excelValue && excelValue.startsWith(qrValue)) {
        const cellAddress = getExcelCellAddress(i + 1, dateColumnIndex + 1);
        worksheet[cellAddress] = { v: 'P' }; // Mark "P" in the correct cell

        matched = true;
        break;
      }
    }

    if (matched) {
      XLSX.writeFile(workbook, uploadedFilePath);

      return res.json({
        message: `Updated attendance for ${qrSubstring} on ${attendanceDate}`,
        downloadLink: `/uploads/${fileName}`
      });
    } else {
      return res.json({ message: 'No match found for QR code.' });
    }
  } catch (error) {
    res.status(500).json({ message: 'Error updating attendance.', error });
  }
});

// Start the server
app.listen(3000, () => {
  console.log('Server started on port 3000');
});
