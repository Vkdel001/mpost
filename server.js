const express = require('express');
const bodyParser = require('body-parser');
const { spawn } = require('child_process');
require('dotenv').config();
const path = require('path');

const app = express();
app.use(bodyParser.text({ type: '*/*' }));

// 🔥 Serve Excel file from /public as /files
app.use('/files', express.static(path.join(__dirname, 'public')));

app.post('/webhook', (req, res) => {
    const python = spawn('python3', ['process_and_email.py']); // or 'python' depending on environment
    python.stdin.write(req.body);
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
        const fileUrl = `${req.protocol}://${req.get('host')}/files/report.xlsx`;
        res.send({
            message: '✅ Excel report created successfully.',
            download_url: fileUrl
        });
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Webhook running on port ${PORT}`);
});
