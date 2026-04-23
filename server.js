const express = require('express');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const bcrypt = require('bcrypt');

const app = express();
const PORT = 3000;
const FILE_NAME = path.join(__dirname, 'UserCredentials.xlsx');
const SHEET_NAME = 'Users';

app.use(bodyParser.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(__dirname));

async function getWorkbookAndSheet() {
    const workbook = new ExcelJS.Workbook();
    let worksheet;

    if (fs.existsSync(FILE_NAME)) {
        await workbook.xlsx.readFile(FILE_NAME);
        worksheet = workbook.getWorksheet(SHEET_NAME);
        if (!worksheet) {
            worksheet = workbook.addWorksheet(SHEET_NAME);
        }
    } else {
        worksheet = workbook.addWorksheet(SHEET_NAME);
    }

    if (worksheet.rowCount === 0) {
        worksheet.columns = [
            { header: 'Username', key: 'username', width: 25 },
            { header: 'Hashed Password', key: 'password', width: 70 },
            { header: 'Security Question', key: 'securityQuestion', width: 35 },
            { header: 'Security Answer', key: 'securityAnswer', width: 35 }
        ];
    }

    return { workbook, worksheet };
}

function normalizeCell(cellValue) {
    if (cellValue === null || cellValue === undefined) return '';
    if (typeof cellValue === 'object' && cellValue.text) return String(cellValue.text).trim();
    return String(cellValue).trim();
}

function getUserFromRow(row) {
    return {
        username: normalizeCell(row.getCell(1).value),
        password: normalizeCell(row.getCell(2).value),
        securityQuestion: normalizeCell(row.getCell(3).value),
        securityAnswer: normalizeCell(row.getCell(4).value)
    };
}

function findUserRowByUsername(worksheet, username) {
    const target = String(username || '').trim().toLowerCase();
    for (let i = 2; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        const rowUsername = normalizeCell(row.getCell(1).value).toLowerCase();
        if (rowUsername && rowUsername === target) {
            return row;
        }
    }
    return null;
}

app.post('/api/users', async (req, res) => {
    try {
        const { username, password, securityQuestion, securityAnswer } = req.body;

        if (!username || !password || !securityQuestion || !securityAnswer) {
            return res.status(400).json({ message: 'username, password, securityQuestion, and securityAnswer are required.' });
        }

        const { workbook, worksheet } = await getWorkbookAndSheet();
        const existingRow = findUserRowByUsername(worksheet, username);
        if (existingRow) {
            return res.status(409).json({ message: 'Username already exists.' });
        }

        const hashedPassword = await bcrypt.hash(password, 10);
        const normalizedAnswer = String(securityAnswer).trim().toLowerCase();

        worksheet.addRow({
            username: String(username).trim(),
            password: hashedPassword,
            securityQuestion: String(securityQuestion).trim(),
            securityAnswer: normalizedAnswer
        });

        await workbook.xlsx.writeFile(FILE_NAME);
        return res.status(201).json({ message: 'User created successfully.' });
    } catch (error) {
        console.error('POST /api/users error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.get('/api/users', async (req, res) => {
    try {
        const { worksheet } = await getWorkbookAndSheet();
        const users = [];

        for (let i = 2; i <= worksheet.rowCount; i++) {
            const row = worksheet.getRow(i);
            const user = getUserFromRow(row);
            if (!user.username) continue;
            users.push({
                username: user.username,
                securityQuestion: user.securityQuestion
            });
        }

        return res.json(users);
    } catch (error) {
        console.error('GET /api/users error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.get('/api/users/:username', async (req, res) => {
    try {
        const { worksheet } = await getWorkbookAndSheet();
        const row = findUserRowByUsername(worksheet, req.params.username);

        if (!row) {
            return res.status(404).json({ message: 'User not found.' });
        }

        const user = getUserFromRow(row);
        return res.json({
            username: user.username,
            securityQuestion: user.securityQuestion
        });
    } catch (error) {
        console.error('GET /api/users/:username error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.put('/api/users/:username', async (req, res) => {
    try {
        const { workbook, worksheet } = await getWorkbookAndSheet();
        const row = findUserRowByUsername(worksheet, req.params.username);

        if (!row) {
            return res.status(404).json({ message: 'User not found.' });
        }

        const updates = req.body || {};
        const existing = getUserFromRow(row);

        let nextUsername = existing.username;
        let nextPasswordHash = existing.password;
        let nextSecurityQuestion = existing.securityQuestion;
        let nextSecurityAnswer = existing.securityAnswer;

        if (updates.username && String(updates.username).trim().toLowerCase() !== existing.username.toLowerCase()) {
            const duplicate = findUserRowByUsername(worksheet, updates.username);
            if (duplicate) {
                return res.status(409).json({ message: 'Target username already exists.' });
            }
            nextUsername = String(updates.username).trim();
        }

        if (updates.password) {
            nextPasswordHash = await bcrypt.hash(String(updates.password), 10);
        }

        if (updates.securityQuestion) {
            nextSecurityQuestion = String(updates.securityQuestion).trim();
        }

        if (updates.securityAnswer) {
            nextSecurityAnswer = String(updates.securityAnswer).trim().toLowerCase();
        }

        row.getCell(1).value = nextUsername;
        row.getCell(2).value = nextPasswordHash;
        row.getCell(3).value = nextSecurityQuestion;
        row.getCell(4).value = nextSecurityAnswer;
        row.commit();

        await workbook.xlsx.writeFile(FILE_NAME);
        return res.json({ message: 'User updated successfully.' });
    } catch (error) {
        console.error('PUT /api/users/:username error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.delete('/api/users/:username', async (req, res) => {
    try {
        const { workbook, worksheet } = await getWorkbookAndSheet();
        const row = findUserRowByUsername(worksheet, req.params.username);

        if (!row) {
            return res.status(404).json({ message: 'User not found.' });
        }

        worksheet.spliceRows(row.number, 1);
        await workbook.xlsx.writeFile(FILE_NAME);

        return res.json({ message: 'User deleted successfully.' });
    } catch (error) {
        console.error('DELETE /api/users/:username error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.post('/api/auth/login', async (req, res) => {
    try {
        const { username, password } = req.body;

        if (!username || !password) {
            return res.status(400).json({ message: 'username and password are required.' });
        }

        const { worksheet } = await getWorkbookAndSheet();
        const row = findUserRowByUsername(worksheet, username);

        if (!row) {
            return res.status(404).json({ message: 'User not found.' });
        }

        const user = getUserFromRow(row);
        const matched = await bcrypt.compare(String(password), user.password);

        if (!matched) {
            return res.status(401).json({ message: 'Incorrect password.' });
        }

        return res.json({ message: 'Login successful.', username: user.username });
    } catch (error) {
        console.error('POST /api/auth/login error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.post('/api/auth/forgot-password-reset', async (req, res) => {
    try {
        const { username, securityQuestion, securityAnswer, newPassword } = req.body;

        if (!username || !securityQuestion || !securityAnswer || !newPassword) {
            return res.status(400).json({ message: 'username, securityQuestion, securityAnswer, and newPassword are required.' });
        }

        const { workbook, worksheet } = await getWorkbookAndSheet();
        const row = findUserRowByUsername(worksheet, username);

        if (!row) {
            return res.status(404).json({ message: 'Username not found.' });
        }

        const user = getUserFromRow(row);
        const normalizedAnswer = String(securityAnswer).trim().toLowerCase();

        if (
            user.securityQuestion !== String(securityQuestion).trim() ||
            user.securityAnswer !== normalizedAnswer
        ) {
            return res.status(401).json({ message: 'Incorrect security question or answer.' });
        }

        const hashedPassword = await bcrypt.hash(String(newPassword), 10);
        row.getCell(2).value = hashedPassword;
        row.commit();

        await workbook.xlsx.writeFile(FILE_NAME);
        return res.json({ message: 'Password reset successfully.' });
    } catch (error) {
        console.error('POST /api/auth/forgot-password-reset error:', error);
        return res.status(500).json({ message: 'Internal server error.' });
    }
});

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'Landing.html'));
});

app.listen(PORT, () => {
    console.log(`Server is running at http://localhost:${PORT}`);
});
