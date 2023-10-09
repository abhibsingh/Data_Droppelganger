const express = require('express');
const fileUpload = require('express-fileupload');
const ExcelJS = require('exceljs');
const cors = require("cors");
const axios = require('axios');

const app = express();
app.use(express.json());
app.use(fileUpload());
app.use(cors())

const OPENAI_API_URL = 'https://api.openai.com/v1/chat/completions';
const RANDOMUSER_API_URL = 'https://randomuser.me/api/?results=100';
const API_KEY = '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'; // Replace with your OpenAI API Key

app.post('/generate', async (req, res) => {
    const file = req.files.sampleFile;
    const sampleData = file.data.toString('utf8').split('\n').map(line => line.split(','));
    const headers = sampleData[0];

    try {
        // Fetch data using OpenAI
        const openaiResponse = await axios.post(OPENAI_API_URL, {
            model: "gpt-4",
            messages:[{
                role: "user",
                content: `
    Create a five-column CSV. Generate data as follows:
    "Title": Select from Mr., Mrs, Dr, Miss, or Prof.
    "Name": Generate a random English name.
    "Email Address": Generate a random email in the format "name@domain.com".
    "Task": Randomly select between "complete" or "incomplete".
    "Date": Generate a random date in the format DD-MM-YYYY.
    `
            }],
            temperature: 1,
            max_tokens: 256,
            top_p: 1,
            frequency_penalty: 0,
            presence_penalty: 0
        }, {
            headers: {
                'Authorization': `Bearer ${API_KEY}`,
                'User-Agent': 'OpenAI-Node-Client',
            }
        });
        
        let openaiDataText = "";
        if (openaiResponse.data.choices && openaiResponse.data.choices[0].message && openaiResponse.data.choices[0].message.content) {
            openaiDataText = openaiResponse.data.choices[0].message.content.trim();
        }
        const openaiData = openaiDataText.split('\n').map(line => line.split(','));

        // Fetch random users
        const randomUsersResponse = await axios.get(RANDOMUSER_API_URL);
        const randomUsers = randomUsersResponse.data.results;
        const randomUserData = randomUsers.map((user) => {
            return [
                ["Mr.", "Mrs.", "Dr.", "Miss", "Prof."][Math.floor(Math.random() * 5)],
                `${user.name.first} ${user.name.last}`,
                user.email,
                ["complete", "incomplete"][Math.floor(Math.random() * 2)],
                new Date(user.dob.date).toLocaleDateString('en-GB') // Format: DD/MM/YYYY
            ];
        });

        // Combine OpenAI data with RandomUser data
        const combinedData = [...openaiData, ...randomUserData];

        // Create an Excel workbook and worksheet
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Generated Data');

        // Add headers to the worksheet
        worksheet.addRow(headers);

        // Add combined data to the worksheet
        combinedData.forEach(row => worksheet.addRow(row));

        // Write the Excel file to a buffer
        const buffer = await workbook.xlsx.writeBuffer();

        // Set the response headers and send the buffer
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=out.xlsx');
        res.send(buffer);

    } catch (error) {
        console.error('Error:', error.response ? error.response.data : error.message);
        res.status(500).send('Internal Server Error');
    }
});

app.listen(4028, () => {
    console.log('Server is running on port 4028');
});
