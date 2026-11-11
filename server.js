const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const { exec } = require('child_process');

const app = express();
const PORT = process.env.PORT || 3000;
const EXCEL_FILE_PATH = path.join(__dirname, 'records.xlsx');
const NIKSHAY_EXCEL_FILE_PATH = path.join(__dirname, 'nikshay_records.xlsx');

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// Generate / update Excel file with provided batch rows
app.post('/api/generate-excel', async (req, res) => {
  try {
    const { rows } = req.body;
    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ success: false, message: 'No rows provided to generate Excel file.' });
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Voice Records App';
    workbook.created = new Date();

    const sheet = workbook.addWorksheet('Voice Records', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    // Define columns
    sheet.columns = [
      { header: '#', key: 'id', width: 8 },
      { header: 'Entry Time', key: 'timestamp', width: 22 },
      { header: 'First Name', key: 'firstName', width: 18 },
      { header: 'Last Name', key: 'lastName', width: 18 },
      { header: 'Age', key: 'age', width: 10 },
      { header: 'Sex', key: 'sex', width: 12 },
      { header: 'Phone No', key: 'phoneNo', width: 20 },
      { header: 'Address', key: 'address', width: 38 }
    ];

    // Style the header row
    const headerRow = sheet.getRow(1);
    headerRow.height = 30;
    headerRow.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF1E293B' } // Slate 800
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    // Border style
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    // Add rows
    rows.forEach((row, index) => {
      const addedRow = sheet.addRow({
        id: index + 1,
        timestamp: row.timestamp || new Date().toLocaleString(),
        firstName: row.firstName || '',
        lastName: row.lastName || '',
        age: row.age ? Number(row.age) || row.age : '',
        sex: row.sex || '',
        phoneNo: row.phoneNo || '',
        address: row.address || ''
      });

      addedRow.height = 24;
      addedRow.font = { name: 'Segoe UI', size: 10.5 };
      addedRow.alignment = { vertical: 'middle', horizontal: 'left' };

      // Center specific columns
      addedRow.getCell('id').alignment = { vertical: 'middle', horizontal: 'center' };
      addedRow.getCell('age').alignment = { vertical: 'middle', horizontal: 'center' };
      addedRow.getCell('sex').alignment = { vertical: 'middle', horizontal: 'center' };
      addedRow.getCell('phoneNo').alignment = { vertical: 'middle', horizontal: 'center' };

      // Zebra striping
      if (index % 2 === 1) {
        addedRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF8FAFC' } // Slate 50
        };
      }

      // Add borders
      addedRow.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = thinBorder;
      });
    });

    // Write file to workspace
    await workbook.xlsx.writeFile(EXCEL_FILE_PATH);

    res.json({
      success: true,
      message: `Successfully generated Excel with ${rows.length} records.`,
      filename: 'records.xlsx',
      count: rows.length,
      fullPath: EXCEL_FILE_PATH
    });
  } catch (error) {
    console.error('Error generating Excel file:', error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Download Excel file
app.get('/api/download', (req, res) => {
  if (fs.existsSync(EXCEL_FILE_PATH)) {
    res.download(EXCEL_FILE_PATH, 'records.xlsx');
  } else {
    res.status(404).json({ success: false, message: 'No records.xlsx file found. Please generate one first.' });
  }
});

// Generate / update Excel file for Nikshay IDs batch
app.post('/api/generate-nikshay-excel', async (req, res) => {
  try {
    const { rows } = req.body;
    if (!rows || !Array.isArray(rows) || rows.length === 0) {
      return res.status(400).json({ success: false, message: 'No Nikshay IDs provided to generate Excel file.' });
    }

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Nikshay IDs Voice App';
    workbook.created = new Date();

    const sheet = workbook.addWorksheet('Nikshay IDs', {
      views: [{ state: 'frozen', ySplit: 1 }]
    });

    // Define columns
    sheet.columns = [
      { header: '#', key: 'id', width: 8 },
      { header: 'Entry Time', key: 'timestamp', width: 24 },
      { header: 'Nikshay ID', key: 'nikshayId', width: 28 }
    ];

    // Style the header row
    const headerRow = sheet.getRow(1);
    headerRow.height = 30;
    headerRow.font = { name: 'Segoe UI', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    headerRow.fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: { argb: 'FF0F766E' } // Teal 700
    };
    headerRow.alignment = { vertical: 'middle', horizontal: 'center' };

    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right: { style: 'thin', color: { argb: 'FFE2E8F0' } }
    };

    // Add rows
    rows.forEach((row, index) => {
      const addedRow = sheet.addRow({
        id: index + 1,
        timestamp: row.timestamp || new Date().toLocaleString(),
        nikshayId: row.nikshayId || row.id || ''
      });

      addedRow.height = 24;
      addedRow.font = { name: 'Segoe UI', size: 11 };
      addedRow.alignment = { vertical: 'middle', horizontal: 'center' };

      // Zebra striping
      if (index % 2 === 1) {
        addedRow.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: { argb: 'FFF0FDFA' } // Teal 50
        };
      }

      addedRow.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = thinBorder;
      });
    });

    await workbook.xlsx.writeFile(NIKSHAY_EXCEL_FILE_PATH);

    res.json({
      success: true,
      message: `Successfully generated Excel with ${rows.length} Nikshay IDs.`,
      filename: 'nikshay_records.xlsx',
      count: rows.length,
      fullPath: NIKSHAY_EXCEL_FILE_PATH
    });
  } catch (error) {
    console.error('Error generating Nikshay Excel file:', error);
    res.status(500).json({ success: false, message: error.message });
  }
});

// Download Nikshay Excel file
app.get('/api/download-nikshay', (req, res) => {
  if (fs.existsSync(NIKSHAY_EXCEL_FILE_PATH)) {
    res.download(NIKSHAY_EXCEL_FILE_PATH, 'nikshay_records.xlsx');
  } else {
    res.status(404).json({ success: false, message: 'No nikshay_records.xlsx file found. Please generate one first.' });
  }
});

// Get existing Nikshay records
app.get('/api/nikshay-records', async (req, res) => {
  try {
    if (!fs.existsSync(NIKSHAY_EXCEL_FILE_PATH)) {
      return res.json({ success: true, rows: [] });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(NIKSHAY_EXCEL_FILE_PATH);
    const sheet = workbook.getWorksheet('Nikshay IDs') || workbook.worksheets[0];
    if (!sheet) {
      return res.json({ success: true, rows: [] });
    }

    const rows = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return;
      rows.push({
        id: row.getCell(1).value,
        timestamp: row.getCell(2).value,
        nikshayId: String(row.getCell(3).value || '')
      });
    });

    res.json({ success: true, rows });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

// Open folder in Windows Explorer
app.post('/api/open-folder', (req, res) => {
  exec(`explorer.exe "${__dirname}"`, (err) => {
    if (err) {
      return res.status(500).json({ success: false, message: 'Could not open folder in Explorer' });
    }
    res.json({ success: true, message: 'Opened folder in File Explorer' });
  });
});

// Get existing records if any
app.get('/api/records', async (req, res) => {
  try {
    if (!fs.existsSync(EXCEL_FILE_PATH)) {
      return res.json({ success: true, rows: [] });
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(EXCEL_FILE_PATH);
    const sheet = workbook.getWorksheet('Voice Records') || workbook.worksheets[0];
    if (!sheet) {
      return res.json({ success: true, rows: [] });
    }

    const rows = [];
    sheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header
      rows.push({
        id: row.getCell(1).value,
        timestamp: row.getCell(2).value,
        firstName: row.getCell(3).value,
        lastName: row.getCell(4).value,
        age: row.getCell(5).value,
        sex: row.getCell(6).value,
        phoneNo: row.getCell(7).value,
        address: row.getCell(8).value
      });
    });

    res.json({ success: true, rows });
  } catch (error) {
    res.status(500).json({ success: false, message: error.message });
  }
});

const os = require('os');

function getNetworkIp() {
  const interfaces = os.networkInterfaces();
  for (const name of Object.keys(interfaces)) {
    for (const iface of interfaces[name]) {
      if (iface.family === 'IPv4' && !iface.internal) {
        return iface.address;
      }
    }
  }
  return '172.20.10.2';
}

const HOST = '0.0.0.0';
app.listen(PORT, HOST, () => {
  const networkIp = getNetworkIp();
  console.log(`Voice Form & Excel Server running on:`);
  console.log(`  > Local:   http://localhost:${PORT}`);
  console.log(`  > Network: http://${networkIp}:${PORT}`);
});
