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

// Notion integration token and database ID
const NOTION_TOKEN = 'ntn_WA6421894517rRDIKWaOMPKuThVcFADzI6BscA5lDmc5H1';
const NOTION_DB_ID = '11202647088b806b9d7dc3836f8aa335'; // or '11202647-088b-806b-9d7d-c3836f8aa335'

app.use(cors({
    origin: 'https://ggpcapplicationform.netlify.app'
}));
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
        const course = req.body.course;
        const expectedSalary = req.body.expectedSalary;
        const availability = req.body.availability;

        console.log('Received birthdate:', birthdate);
        console.log('Received position:', position);

        // 1. Upload file to Kintone
        const formData = new FormData();
        formData.append('file', file.buffer, file.originalname);

        const uploadRes = await fetch(`https://${KINTONE_DOMAIN}/k/v1/file.json`, {
            method: 'POST',
            headers: {
                'X-Cybozu-API-Token': KINTONE_API_TOKEN
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
                Expected_Salary: { value: expectedSalary },
                Availability: { value: availability }
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
                    Expected_Salary: { value: expectedSalary },
                    Availability: { value: availability }
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

        const recordId = recordRes.data.id;

        console.log(`Created Kintone record: ${recordId}`);

        // 3. Send confirmation email via Netlify
        const emailRes = await axios.post(
            'https://confirmation-email-backend.netlify.app/.netlify/functions/send-email',
            {
                name,
                email,
                recordId
            }
        );

        console.log(`Email sent:`, emailRes.data);

        res.json({ success: true });
    } catch (err) {
        console.error('Server error:', err);
        res.status(500).json({ error: err.message });
    }
});

app.post('/check-name', async (req, res) => {
    try {
        const { name, email, position } = req.body;
        // Query Kintone for records with the same name
        const query = `Full_Name = "${name}" or Email = "${email}"`;
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
        let exists = false;
        let sameEmailSamePosition = false;
        let sameEmailDifferentPosition = false;
        if (response.data.records && response.data.records.length > 0) {
            for (const rec of response.data.records) {
                const recEmail = rec.Email?.value || '';
                const recPosition = rec.Position?.value || '';
                if (recEmail === email && recPosition === position) {
                    sameEmailSamePosition = true;
                    exists = true;
                } else if (recEmail === email && recPosition !== position) {
                    sameEmailDifferentPosition = true;
                }
            }
        }
        return res.json({ exists, sameEmailSamePosition, sameEmailDifferentPosition });
    } catch (err) {
        console.error('Error checking name/email:', err);
        res.status(500).json({ error: 'Server error' });
    }
});

app.get('/health', (req, res) => {
    res.send('OK');
});

// Temporary endpoint to inspect Notion DB schema
app.get('/notion-schema', async (req, res) => {
    try {
        const notionRes = await axios.get(
            `https://api.notion.com/v1/databases/${NOTION_DB_ID}`,
            {
                headers: {
                    'Authorization': `Bearer ${NOTION_TOKEN}`,
                    'Notion-Version': '2022-06-28',
                    'Content-Type': 'application/json'
                }
            }
        );
        res.json(notionRes.data);
    } catch (err) {
        // Log the entire error object for debugging
        console.error('Full error object:', err);
        // Return all possible error info
        res.status(500).json({
            error: 'Failed to fetch Notion schema.',
            details: err.response?.data || null,
            message: err.message,
            stack: err.stack,
            code: err.code || null,
            config: err.config || null
        });
    }
});

// New endpoint to get job positions from Notion
app.get('/get-jobs', async (req, res) => {
    try {
        const response = await axios.post(
            `https://api.notion.com/v1/databases/${NOTION_DB_ID}/query`,
            {
                // We add a filter to only get pages where the "Status" is "Open"
                filter: {
                    property: 'Employment Status', // Make sure you have a 'Status' column in Notion
                    select: {
                        equals: 'Open' // And you have an option called 'Open'
                    }
                },
                // We sort by the "Position Name" alphabetically
                sorts: [
                    {
                        property: 'Position Name', // 
                        direction: 'ascending'
                    }
                ]
            },
            {
                headers: {
                    'Authorization': `Bearer ${NOTION_TOKEN}`,
                    'Notion-Version': '2022-06-28',
                    'Content-Type': 'application/json'
                }
            }
        );

        // Extract the job titles from Notion's response
        const jobs = response.data.results.map(page => {
            // Assumes your job title column is named 'Job Title'
            return page.properties['Position Name']?.title[0]?.plain_text;
        }).filter(Boolean); // Filter out any empty results

        res.json(jobs);

    } catch (err) {
        console.error('Error fetching jobs from Notion:', err.response ? err.response.data : err.message);
        res.status(500).json({ error: 'Failed to fetch jobs from Notion.' });
    }
});

app.listen(3000, () => {
    console.log('Proxy server running on https://ggpcapplicationform.onrender.com');
});