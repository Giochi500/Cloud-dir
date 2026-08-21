const express = require('express');
const XLSX = require('xlsx');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const { Pool } = require('pg');
const jwt = require('jsonwebtoken');
const bcrypt = require('bcryptjs');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

const SALARY_DIR = 'C:\\Users\\Legion\\OneDrive\\დოკუმენტები\\My app\\data';

// PostgreSQL connection
const pool = new Pool({
  connectionString: process.env.DATABASE_URL,
  ssl: { rejectUnauthorized: false }
});

// Initialize database
async function initDatabase() {
  try {
    await pool.query(`
      CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        username VARCHAR(255) UNIQUE NOT NULL,
        email VARCHAR(255) UNIQUE NOT NULL,
        password VARCHAR(255) NOT NULL,
        role VARCHAR(50) NOT NULL CHECK (role IN ('admin', 'manager', 'hr')),
        department VARCHAR(255),
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
      )
    `);
    console.log('Database initialized');
  } catch (err) {
    console.error('Error initializing database:', err);
  }
}

initDatabase();

// JWT middleware
function verifyToken(req, res, next) {
  const token = req.headers['authorization']?.split(' ')[1];
  if (!token) return res.status(401).json({ error: 'No token provided' });

  jwt.verify(token, process.env.JWT_SECRET, (err, decoded) => {
    if (err) return res.status(401).json({ error: 'Invalid token' });
    req.user = decoded;
    next();
  });
}

// Role check middleware
function checkRole(roles) {
  return (req, res, next) => {
    if (!roles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Access denied' });
    }
    next();
  };
}

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

// POST: Register new user
app.post('/api/register', async (req, res) => {
  try {
    const { username, email, password, role, department } = req.body;

    if (!username || !email || !password || !role) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const hashedPassword = await bcrypt.hash(password, 10);

    const result = await pool.query(
      'INSERT INTO users (username, email, password, role, department) VALUES ($1, $2, $3, $4, $5) RETURNING id, username, email, role',
      [username, email, hashedPassword, role, department || null]
    );

    res.json({ success: true, user: result.rows[0] });
  } catch (err) {
    if (err.code === '23505') {
      return res.status(400).json({ error: 'Username or email already exists' });
    }
    res.status(500).json({ error: err.message });
  }
});

// POST: Login
app.post('/api/login', async (req, res) => {
  try {
    const { username, password } = req.body;

    if (!username || !password) {
      return res.status(400).json({ error: 'Username and password required' });
    }

    const result = await pool.query('SELECT * FROM users WHERE username = $1', [username]);

    if (result.rows.length === 0) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    const user = result.rows[0];
    const validPassword = await bcrypt.compare(password, user.password);

    if (!validPassword) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }

    const token = jwt.sign(
      { id: user.id, username: user.username, role: user.role, department: user.department },
      process.env.JWT_SECRET,
      { expiresIn: '24h' }
    );

    res.json({ success: true, token, user: { id: user.id, username: user.username, role: user.role } });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// GET: Current user
app.get('/api/user', (req, res) => {
  res.json({ user: req.user });
});

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

// POST: Add new employee (admin, hr only)
app.post('/api/add-employee/:month', (req, res) => {
  try {
    const { month } = req.params;
    const { name, position, location, department, personalId, contractType, salary } = req.body;

    const files = getSalaryFiles();
    if (!files[month]) {
      return res.status(404).json({ error: `No data for month ${month}` });
    }

    const filePath = files[month];
    const data = readExcelFile(filePath);

    // Create new employee object matching Excel structure
    const newEmployee = {
      '__EMPTY': department,
      '__EMPTY_1': 'შპს მეგა ჰოლდინგი',
      '__EMPTY_2': contractType,
      '__EMPTY_3': location,
      '__EMPTY_4': name,
      '__EMPTY_5': position,
      '__EMPTY_6': personalId,
      '__EMPTY_7': 'საპენსიოს გარეშე',
      '__EMPTY_8': 0,
      '__EMPTY_9': 0,
      '__EMPTY_10': salary,
      '__EMPTY_11': 0,
      '__EMPTY_12': 0,
      '__EMPTY_13': 0,
      '__EMPTY_14': 0,
      '__EMPTY_15': 0,
      '__EMPTY_16': salary,
      '__EMPTY_17': 0,
      '__EMPTY_18': 0,
      '__EMPTY_19': 0,
      '__EMPTY_20': 0,
      '__EMPTY_21': 0,
      '__EMPTY_22': 0,
      '__EMPTY_23': '',
      '__EMPTY_24': ''
    };

    data.push(newEmployee);

    // Get headers
    const headers = data.length > 0 ? Object.keys(data[0]) : [];

    // Write back to Excel
    const rows = data.map(emp => headers.map(h => emp[h]));
    const workbook = XLSX.readFile(filePath);
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    workbook.Sheets[workbook.SheetNames[0]] = worksheet;
    XLSX.writeFile(workbook, filePath);

    res.json({ success: true, message: 'Employee added successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// DELETE: Delete employee (admin, hr only)
app.delete('/api/delete-employee/:month/:employeeId', (req, res) => {
  try {
    const { month, employeeId } = req.params;
    const empIdx = parseInt(employeeId) - 1;

    const files = getSalaryFiles();
    if (!files[month]) {
      return res.status(404).json({ error: `No data for month ${month}` });
    }

    const filePath = files[month];
    const data = readExcelFile(filePath);

    if (empIdx < 0 || empIdx >= data.length) {
      return res.status(404).json({ error: 'Employee not found' });
    }

    // Remove employee from array
    data.splice(empIdx, 1);

    // Get headers
    const headers = data.length > 0 ? Object.keys(data[0]) : [];

    // Write back to Excel
    const rows = data.map(emp => headers.map(h => emp[h]));
    const workbook = XLSX.readFile(filePath);
    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    workbook.Sheets[workbook.SheetNames[0]] = worksheet;
    XLSX.writeFile(workbook, filePath);

    res.json({ success: true, message: 'Employee deleted successfully' });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// PUT: Update salary entry for employee (admin, hr only)
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
  res.sendFile(path.join(__dirname, 'public', 'dashboard.html'));
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
  console.log(`Salary directory: ${SALARY_DIR}`);
});
