const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const SALARY_DIR = 'C:\\Users\\Legion\\OneDrive\\Desktop\\ხელფასებისთვის';

// Helper: Get all salary files by month
function getSalaryFiles() {
  const files = {};
  try {
    const months = fs.readdirSync(SALARY_DIR);
    months.forEach(month => {
      const monthPath = path.join(SALARY_DIR, month);
      if (fs.statSync(monthPath).isDirectory()) {
        const monthFiles = fs.readdirSync(monthPath);
        monthFiles.forEach(file => {
          if (file.includes('სტანდარტიზებული') && file.endsWith('.xlsx')) {
            files[month] = path.join(monthPath, file);
          }
        });
      }
    });
  } catch (err) {
    console.error('Error reading salary directory:', err);
  }
  return files;
}

// Helper: Read Excel file
function readExcelFile(filePath) {
  try {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const data = XLSX.utils.sheet_to_json(worksheet);

    // Filter out rows with no employee name (header/empty rows)
    const filtered = data.filter(row => {
      const name = row['__EMPTY_4'];
      const isHeader = name === 'სახელი და გვარი' || name === 'Name' || !name;
      const isComment = String(name).includes('ერთიანი') || String(name).includes('ავტომატური');
      return name && !isHeader && !isComment && name !== 'N/A' && String(name).trim() !== '';
    });

    return filtered;
  } catch (err) {
    console.error('Error reading Excel:', err);
    return [];
  }
}

// Helper: Write Excel file
function writeExcelFile(filePath, data, headers) {
  try {
    const workbook = XLSX.readFile(filePath);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Update rows
    const ws_data = [headers, ...data];
    const new_worksheet = XLSX.utils.aoa_to_sheet(ws_data);

    workbook.Sheets[workbook.SheetNames[0]] = new_worksheet;
    XLSX.writeFile(workbook, filePath);
    return true;
  } catch (err) {
    console.error('Error writing Excel:', err);
    return false;
  }
}

// GET: List all employees
app.get('/api/employees', (req, res) => {
  try {
    const files = getSalaryFiles();
    const monthKey = Object.keys(files)[0];

    if (!monthKey) {
      return res.status(404).json({ error: 'No salary files found' });
    }

    const employees = readExcelFile(files[monthKey]);

    // Log first employee to debug column names
    if (employees.length > 0) {
      console.log('First employee keys:', Object.keys(employees[0]));
      console.log('First employee:', employees[0]);
    }

    const employeeList = employees.map((emp, idx) => ({
      id: idx + 1,
      ...emp
    }));

    res.json(employeeList);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET: Get salary data for specific month
app.get('/api/salary/:month', (req, res) => {
  try {
    const month = req.params.month; // e.g., "2026.06"
    const files = getSalaryFiles();

    if (!files[month]) {
      return res.status(404).json({ error: `No data for month ${month}` });
    }

    const data = readExcelFile(files[month]);
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET: Get available months
app.get('/api/months', (req, res) => {
  try {
    const files = getSalaryFiles();
    const months = Object.keys(files).sort();
    res.json(months);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET: Get unique positions/titles
app.get('/api/positions', (req, res) => {
  try {
    const files = getSalaryFiles();
    const monthKey = Object.keys(files)[0];

    if (!monthKey) {
      return res.status(404).json({ error: 'No salary files found' });
    }

    const employees = readExcelFile(files[monthKey]);
    const positions = [...new Set(employees.map(emp => emp['__EMPTY_5']).filter(Boolean))].sort();
    res.json(positions);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// PUT: Update salary entry for employee
app.put('/api/salary/:month/:employeeId', (req, res) => {
  try {
    const { month, employeeId } = req.params;
    const { ავანსი, გამომუშავებით, მოჭრილი } = req.body;

    const files = getSalaryFiles();
    if (!files[month]) {
      return res.status(404).json({ error: `No data for month ${month}` });
    }

    const filePath = files[month];
    const data = readExcelFile(filePath);

    const empIdx = parseInt(employeeId) - 1;
    if (empIdx < 0 || empIdx >= data.length) {
      return res.status(404).json({ error: 'Employee not found' });
    }

    // Update fields
    if (ავანსი !== undefined) data[empIdx]['ავანსი (Net)'] = ავანსი;
    if (გამომუშავებით !== undefined) data[empIdx]['გამომუშავებით (Net)'] = გამომუშავებით;
    if (მოჭრილი !== undefined) data[empIdx]['მოჭრილი (net)'] = მოჭრილი;

    // Get headers
    const headers = Object.keys(data[0]);

    // Write back
    const rows = data.map(emp => headers.map(h => emp[h]));

    const workbook = XLSX.readFile(filePath);
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    workbook.Sheets[workbook.SheetNames[0]] = worksheet;
    XLSX.writeFile(workbook, filePath);

    res.json({ success: true, message: 'Updated successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Serve index.html for root
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`Salary directory: ${SALARY_DIR}`);
});
