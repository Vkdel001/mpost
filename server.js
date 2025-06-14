const express = require('express');
const bodyParser = require('body-parser');
const { spawn } = require('child_process');
require('dotenv').config();

const app = express();
app.use(bodyParser.text({ type: '*/*' }));

app.post('/webhook', (req, res) => {
    const python = spawn('python3', ['process_and_email.py']);
    let resultOutput = '';

    python.stdin.write(req.body);
    python.stdin.end();

    python.stdout.on('data', (data) => {
        resultOutput += data.toString();
        console.log(`Python output: ${data}`);
    });

    python.stderr.on('data', (data) => {
        console.error(`Python error: ${data}`);
    });

    python.on('close', (code) => {
        if (code === 0) {
            res.status(200).send(`✅ Success:\n${resultOutput}`);
        } else {
            res.status(500).send(`❌ Python script failed with code ${code}\n${resultOutput}`);
        }
    });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`🚀 Webhook running on port ${PORT}`);
});
