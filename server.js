const express = require('express');
const bodyParser = require('body-parser');
const { spawn } = require('child_process');
require('dotenv').config();
const path = require('path');

const app = express();

// Accept JSON input (e.g., { "rawText": "..." })
app.use(bodyParser.json());

// Serve static files like the Excel report
app.use('/files', express.static(path.join(__dirname, 'public')));

app.post('/webhook', (req, res) => {
    const rawText = req.body.rawText;

    if (!rawText) {
        return res.status(400).send({ error: 'Missing rawText in JSON payload' });
    }

    const python = spawn('python3', ['process_and_email.py']);

    python.stdin.write(rawText);
    python.stdin.end();

    let output = '';
    python.stdout.on('data', (data) => {
        output += data.toString();
        console.log(`Python output: ${data}`);
    });

    python.stderr.on('data', (data) => {
        console.error(`Python error: ${data}`);
    });

    python.on('close', (code) => {
        const fileName = output.trim(); // output should be just the file name like "report_20250614_161045.xlsx"
        const fileUrl = `${req.protocol}://${req.get('host')}/files/${fileName}`;
        res.send({
            message: '✅ Excel report created successfully.',
            download_url: fileUrl,
            filename: fileName
        });
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Webhook running on port ${PORT}`);
});
