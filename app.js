// dotenv setup 
if(process.env.NODE_ENV != "production"){
       require("dotenv").config();
}

const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const mysql = require('mysql2/promise');
const fs = require('fs');
const path = require('path');
const bodyParser = require('body-parser');
const app = express();

// ========================
// DATABASE CONNECTION POOL
// ========================


app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json())

app.set("view engine", "ejs");
app.set("Views", path.join(__dirname, "/views"));

const pool = mysql.createPool({
    host: process.env.HOST,
    user: process.env.USER,
    password: process.env.PASSWORD, // Replace with your database password
    database: process.env.DB, // Your database name
    ssl: {
         rejectUnauthorized: false
    }
});

if(pool){
    console.log("Connected to database");
}else{
    console.log("Failed to connect to database");
}
// ========================
// MULTER CONFIGURATION
// ========================
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = 'uploads/';
        if (!fs.existsSync(uploadDir)) {
            fs.mkdirSync(uploadDir); // Create 'uploads/' directory if it doesn't exist
        }
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        const ext = path.extname(file.originalname); // File extension
        cb(null, `${Date.now()}${ext}`); // Unique filename
    },
});
const upload = multer({ storage: storage });

// ========================
// PARSE EXCEL FILE FUNCTION
// ========================
function parseExcel(filePath) {
    console.log(`Parsing Excel file: ${filePath}`);
    const workbook = XLSX.readFile(filePath);
    const sheetName = workbook.SheetNames[0]; // Assume data is in the first sheet
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: null });

    // Convert Excel date serial numbers to actual Date objects or formatted strings
    data.forEach(row => {
        if (row.dob) {
            // Check if the dob field is an Excel date serial number
            const excelDate = row.dob;
            if (typeof excelDate === 'number') {
                // Convert Excel date serial number to JavaScript Date
                const date = new Date((excelDate - 25569) * 86400 * 1000); // Convert to milliseconds
                row.dob = date.toISOString().split('T')[0]; // Format as YYYY-MM-DD
            }
        }
    });

    console.log('Excel file parsed successfully:', data);
    return data; // Returns an array of objects
}



// ========================
// DATABASE INSERT/UPDATE FUNCTION
// ========================
async function insertOrUpdateData(data) {
    console.log('Inserting/updating data in the database...');
    const connection = await pool.getConnection();
    try {
        for (const student of data) {
            const { rollno, name, class: className, branch, gender, dob, prn } = student;

            if (!prn) {
                console.log(`Skipping invalid record: Missing PRN for student:`, student);
                continue;
            }

            // Check if PRN already exists
            const [rows] = await connection.execute(
                'SELECT prn FROM studentsdata WHERE prn = ?',
                [prn]
            );

            if (rows.length > 0) {
                // Update existing record
                console.log(`Updating record for PRN: ${prn}`);
                await connection.execute(
                    `UPDATE studentsdata 
                     SET rollno = ?, name = ?, class = ?, branch = ?, gender = ?, dob = ?
                     WHERE prn = ?`,
                    [rollno, name, className, branch, gender, dob, prn]
                );
            } else {
                // Insert new record
                console.log(`Inserting new record for PRN: ${prn}`);
                await connection.execute(
                    `INSERT INTO studentsdata (rollno, name, class, branch, gender, dob, prn)
                     VALUES (?, ?, ?, ?, ?, ?, ?)`,
                    [rollno, name, className, branch, gender, dob, prn]
                );
            }
        }
        console.log('Database insert/update completed successfully.');
    } catch (err) {
        console.error('Database operation error:', err);
        throw err;
    } finally {
        connection.release();
    }
}

app.get("/", (req, res)=>{
    res.render("home.ejs");
});



// ========================
// UPLOAD EXCEL API ENDPOINT
//========================
app.post('/upload-excel', upload.single('file'), async (req, res) => {
    if (!req.file) {
        console.error('No file uploaded.');
        return res.status(400).send('No file uploaded. Please upload an Excel file.');
    }

    const filePath = req.file.path;
    console.log(`File received: ${filePath}`);

    try {
        // Parse Excel file
        const data = parseExcel(filePath);

        // Validate Excel data
        console.log('Validating Excel data...');
        for (const row of data) {
            if (!row.rollno || !row.name || !row.class || !row.branch || !row.gender || !row.dob || !row.prn) {
                console.error('Invalid row:', row);
                return res.status(400).send('Invalid data format in Excel file.');
            }
        }
        console.log('Excel data validation successful.');

        // Insert or update data
        await insertOrUpdateData(data);

        res.status(200).send('Excel data inserted/updated successfully.');
    } catch (error) {
        console.error('Error processing file:', error.message);
        res.status(500).send('An error occurred while processing the file.');
    } finally {
        // Clean up temporary file
        try {
            fs.unlinkSync(filePath);
            console.log(`Temporary file deleted: ${filePath}`);
        } catch (err) {
            console.error('Error deleting temporary file:', err);
        }
    }
});


// Hardcoded File Upload Endpoint for Testing
// app.post('/upload-excel', async (req, res) => {
//     const filePath = 'D:\\student_records.xlsx'; // Hardcoded file path
//     console.log(`Using hardcoded file path: ${filePath}`);

//     try {
//         // Parse Excel file
//         const data = parseExcel(filePath);

//         // Validate data format
//         console.log('Validating Excel data...');
//         for (const row of data) {
//             if (!row.rollno || !row.name || !row.class || !row.branch || !row.gender || !row.dob || !row.prn) {
//                 console.log('Invalid data row:', row);
//                 return res.status(400).send('Invalid data format in Excel file.');
//             }
//         }
//         console.log('Excel data validation passed.');

//         // Insert or update data in the database
//         await insertOrUpdateData(data);

//         res.status(200).send('Data successfully inserted/updated.');
//     } catch (error) {
//         console.error('Error processing file:', error);
//         res.status(500).send('An error occurred while processing the file.');
//     }
// });



// ========================
// START SERVER
// ========================
const PORT = 3000;
app.listen(PORT, () => {
    console.log(`Server started on http://localhost:${PORT}`);
});
