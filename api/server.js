// Importing required modules
const express = require('express');
const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');
const app = express();
const port = process.env.PORT || 3002;

// Middleware to parse JSON body
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

// Define file path for XLS
const xlsFilePath = path.join(__dirname, 'data.xls');

// Function to read XLS file
const readXlsFile = () => {
  if (fs.existsSync(xlsFilePath)) {
    const workbook = xlsx.readFile(xlsFilePath);
    const sheet_name_list = workbook.SheetNames;
    const worksheet = workbook.Sheets[sheet_name_list[0]];
    return xlsx.utils.sheet_to_json(worksheet);
  }
  return [];
};

// Function to write to XLS file
const writeXlsFile = (data) => {
  const ws = xlsx.utils.json_to_sheet(data);
  const wb = xlsx.utils.book_new();
  xlsx.utils.book_append_sheet(wb, ws, 'Sheet1');
  xlsx.writeFile(wb, xlsFilePath);
};

// Endpoint to get data from XLS file
app.get('/get-xls-data', (req, res) => {
  const data = readXlsFile();
  res.json({ data });
});

// Endpoint to save data to XLS file
app.post('/save-xls-data', (req, res) => {
  const { data } = req.body;
  if (Array.isArray(data)) {
    writeXlsFile(data);
    res.json({ message: 'Data saved to XLS file successfully' });
  } else {
    res.status(400).json({ error: 'Data should be an array' });
  }
});

// Endpoint to update data in XLS file
app.post('/update-xls-data', (req, res) => {
  const { rowIndex, newData } = req.body;
  if (typeof rowIndex !== 'undefined' && newData) {
    let data = readXlsFile();
    if (rowIndex < data.length) {
      data[rowIndex] = { ...data[rowIndex], ...newData };
      writeXlsFile(data);
      res.json({ message: 'Data updated successfully' });
    } else {
      res.status(400).json({ error: 'Row index out of bounds' });
    }
  } else {
    res.status(400).json({ error: 'Invalid input' });
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
