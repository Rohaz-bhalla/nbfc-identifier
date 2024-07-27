const express = require('express');
const mysql = require('mysql2');
const bodyParser = require('body-parser');
const axios = require('axios');
const ExcelJS = require('exceljs');
const path = require('path');

const app = express();
const port = 3000;

// Middleware
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

// Set the view engine to EJS
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Serve static files (e.g., CSS, JS)
app.use(express.static(path.join(__dirname, 'public')));

// MySQL connection
const db = mysql.createConnection({
    host: 'localhost',
    user: 'root', // Your MySQL username
    password: 'password', // Your MySQL password
    database: 'nbfc_db' // Your MySQL database name
});

db.connect(err => {
    if (err) {
        console.error('Error connecting to MySQL:', err);
        return;
    }
    console.log('MySQL connected...');
});

// Render the front page
app.get('/', (req, res) => {
    res.render('frontpage');
});

// Read Excel file and store data in the database
app.post('/upload', async (req, res) => {
    const filePath = req.body.filepath;

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    const worksheet = workbook.getWorksheet(1);

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        if (rowNumber > 1) { // Skip header row
            const regionalOffice = row.getCell(1).value;
            const nbfcName = row.getCell(2).value;
            const address = row.getCell(3).value;
            const emailId = row.getCell(4).value;

            // Web search automation to find the official website
            axios.get(`https://www.googleapis.com/customsearch/v1?q=${nbfcName}&key=YOUR_API_KEY`)
                .then(response => {
                    const website = response.data.items[0].link; // Simplified for demonstration

                    // Insert data into MySQL
                    const sql = 'INSERT INTO nbfc_details (regional_office, nbfc_name, address, email_id, official_website) VALUES (?, ?, ?, ?, ?)';
                    db.query(sql, [regionalOffice, nbfcName, address, emailId, website], (err, result) => {
                        if (err) throw err;
                        console.log('Data inserted:', result.insertId);
                    });
                })
                .catch(error => {
                    console.log('Error finding website:', error);
                });
        }
    });

    res.send('File processed successfully');
});

// Generate output Excel file
app.get('/download', async (req, res) => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('NBFC Details');

    worksheet.columns = [
        { header: 'Regional Office', key: 'regional_office', width: 30 },
        { header: 'NBFC Name', key: 'nbfc_name', width: 30 },
        { header: 'Address', key: 'address', width: 30 },
        { header: 'Email ID', key: 'email_id', width: 30 },
        { header: 'Official Website', key: 'official_website', width: 30 }
    ];

    db.query('SELECT * FROM nbfc_details', (err, results) => {
        if (err) throw err;

        results.forEach(row => {
            worksheet.addRow(row);
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=nbfc_details.xlsx');
        workbook.xlsx.write(res).then(() => {
            res.end();
        });
    });
});

// Test database connection
app.get('/test-db', (req, res) => {
    const sql = 'SELECT * FROM nbfc_details';
    db.query(sql, (err, results) => {
        if (err) {
            res.status(500).send('Database query failed');
            throw err;
        }
        res.json(results);
    });
});

app.listen(port, () => {
    console.log(`Server running on port ${port}`);
});
