const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const { MongoClient } = require('mongodb');

const app = express();
const upload = multer({ dest: 'uploads/' });


function readExcel(filePath) {
  const workbook = XLSX.readFile(filePath);
  const sheetName = workbook.SheetNames[0]; 
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_json(sheet);
}

// MongoDB connection URL
const uri = "mongodb+srv://ym79793:a9ZoKqdFNfAuiA2z@atlasdb.in2rkqn.mongodb.net/atlas?retryWrites=true&w=majority"; 

// Database and Collection names
const dbName = 'atlas';
const collectionName = 'atlasname';

// Function to insert data into MongoDB
async function insertDataIntoMongoDB(data) {
  const client = new MongoClient(uri, { useNewUrlParser: true, useUnifiedTopology: true });

  try {
    await client.connect();
    const database = client.db(dbName);
    const collection = database.collection(collectionName);

   
    const result = await collection.insertMany(data);
    console.log(`${result.insertedCount} documents inserted.`);
  } catch (err) {
    console.error('Error occurred:', err);
  } finally {
    await client.close();
  }
}


app.post('/upload', upload.single('excelFile'), async (req, res) => {
  if (!req.file) {
    return res.status(400).send('No file uploaded.');
  }

  try {
    const filePath = req.file.path;
    const excelData = readExcel(filePath);
    await insertDataIntoMongoDB(excelData);

    res.status(200).send('File uploaded and data inserted into MongoDB.');
  } catch (err) {
    console.error('Error occurred:', err);
    res.status(500).send('Internal server error.');
  }
});


const PORT = 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
