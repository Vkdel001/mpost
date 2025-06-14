const express = require('express');
const bodyParser = require('body-parser');
const { spawn } = require('child_process');
const fs = require('fs');
const path = require('path');
require('dotenv').config();

const app = express();
app.use(bodyParser.text({ type: '*/*' }));

app.post('/webhook', (req, res) => {
    const python = spawn('python3', ['process_and_email.py']);
    python.stdin.write(req.body);
    python.stdin.end();

    let pythonOutput = '';
    let pythonError = '';

    python.stdout.on('data', (data) => {
        pythonOutput += data.toString();
        console.log(`Python output: ${data}`);
    });

    python.stderr.on('data', (data) => {
        pythonError += data.toString();
        console.error(`Python error: ${data}`);
    });

    python.on('close', (code) => {
        if (code !== 0) {
            return res.status(500).send(`Python script failed with code ${code}\n${pythonError}`);
        }

        // Read the generated Excel file
        const filePath = path.join(__dirname, 'report.xlsx');

        fs.readFile(filePath, (err, fileData) => {
            if (err) {
                console.error('Error reading Excel file:', err);
                return res.status(500).send('Failed to read Excel file');
            }

            // Set headers to prompt download
            res.setHeader('Content-Disposition', 'attachment; filename=report.xlsx');
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');

            res.send(fileData);
        });
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Webhook running on port ${PORT}`);
});
