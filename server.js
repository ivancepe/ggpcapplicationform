const express = require('express');
const multer = require('multer');
const cors = require('cors');
const FormData = require('form-data');
const fetch = require('node-fetch');
const axios = require('axios');

const app = express();
const upload = multer();

const KINTONE_DOMAIN = 'vez7o26y38rb.cybozu.com';
const KINTONE_API_TOKEN = 'SNjXj0CGity20DNSsiJgImu2fj0WIEWyeHvVbyHe';
const KINTONE_APP_ID = '1586';

app.use(cors());
app.use(express.json());

app.post('/apply', upload.single('resume'), async (req, res) => {
    try {
        const name = req.body.name;
        const email = req.body.email;
        const phone = req.body.phone;
        const birthdate = req.body.birthdate;
        const position = req.body.position;
        const file = req.file;
        const age = req.body.age;
        const address = req.body.address;
        const education = req.body.education;
        const course = req.body.course; // New line for course
        const expectedSalary = req.body.expectedSalary;
        const availability = req.body.availability; // New line for availability

        console.log('Received birthdate:', birthdate);
        console.log('Received position:', position);

        // 1. Upload file to Kintone
        const formData = new FormData();
        formData.append('file', file.buffer, file.originalname);

        const uploadRes = await fetch(`https://${KINTONE_DOMAIN}/k/v1/file.json`, {
            method: 'POST',
            headers: {
                'X-Cybozu-API-Token': KINTONE_API_TOKEN
                // Do NOT set 'Content-Type' here!
            },
            body: formData
        });

        if (!uploadRes.ok) {
            const errorText = await uploadRes.text();
            console.error('File upload error:', errorText);
            return res.status(500).json({ error: 'Failed to upload file to Kintone.', details: errorText });
        }

        const uploadData = await uploadRes.json();
        const fileKey = uploadData.fileKey;

        // 2. Create record in Kintone
        console.log('Sending JSON:', JSON.stringify({
            app: KINTONE_APP_ID,
            record: {
                Full_Name: { value: name },
                Phone: { value: phone },
                Birthdate: { value: birthdate },
                Age: { value: age },
                Position: { value: position },
                Email: { value: email },
                Address: { value: address },
                Education: { value: education },
                Resume: { value: [{ fileKey }] },
                Course: { value: course },
                Expected_Salary: { value: expectedSalary }, // <-- Add this line (use your actual field code)
                Availability: { value: availability } // Add this to your Kintone record (use your actual field code)
            }
        }));

        const recordRes = await axios.post(
            `https://${KINTONE_DOMAIN}/k/v1/record.json`,
            {
                app: KINTONE_APP_ID,
                record: {
                    Full_Name: { value: name },
                    Phone: { value: phone },
                    Birthdate: { value: birthdate },
                    Age: { value: age },
                    Position: { value: position },
                    Email: { value: email },
                    Address: { value: address },
                    Education: { value: education },
                    Resume: { value: [{ fileKey }] },
                    Course: { value: course },
                    Expected_Salary: { value: expectedSalary }, // <-- Add this line (use your actual field code)
                    Availability: { value: availability } // Add this to your Kintone record (use your actual field code)
                }
            },
            {
                headers: {
                    'X-Cybozu-API-Token': KINTONE_API_TOKEN
                }
            }
        );

        if (recordRes.status !== 200) {
            console.error('Record creation error:', recordRes.data);
            return res.status(500).json({ error: 'Failed to create record in Kintone.', details: recordRes.data });
        }

        res.json({ success: true });
    } catch (err) {
        console.error('Server error:', err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/check-name', async (req, res) => {
    try {
        const { name } = req.body;
        // Query Kintone for records with the same name
        const query = `Full_Name = "${name}"`;
        const response = await axios.get(
            `https://${KINTONE_DOMAIN}/k/v1/records.json`,
            {
                params: {
                    app: KINTONE_APP_ID,
                    query: query
                },
                headers: {
                    'X-Cybozu-API-Token': KINTONE_API_TOKEN
                }
            }
        );
        if (response.data.records && response.data.records.length > 0) {
            return res.json({ exists: true });
        } else {
            return res.json({ exists: false });
        }
    } catch (err) {
        console.error('Error checking name:', err);
        res.status(500).json({ error: 'Server error' });
    }
});

app.listen(3000, () => {
    console.log('Proxy server running on http://localhost:3000');
});

