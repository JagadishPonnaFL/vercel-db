const fs = require('fs');
const path = require('path');
const xlsx = require('xlsx');

const xlsFilePath = path.join(__dirname, '../data.xls');

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

// Default export for Vercel's serverless function
module.exports = async (req, res) => {
  const { method } = req;

  if (method === 'GET' && req.url === '/get-xls-data') {
    const data = readXlsFile();
    return res.status(200).json({ data });
  }

  if (method === 'POST' && req.url === '/save-xls-data') {
    const { data } = req.body;
    if (Array.isArray(data)) {
      writeXlsFile(data);
      return res.status(200).json({ message: 'Data saved to XLS file successfully' });
    } else {
      return res.status(400).json({ error: 'Data should be an array' });
    }
  }

  if (method === 'POST' && req.url === '/update-xls-data') {
    const { rowIndex, newData } = req.body;
    if (rowIndex !== undefined && newData) {
      let data = readXlsFile();
      if (rowIndex < data.length) {
        data[rowIndex] = { ...data[rowIndex], ...newData };
        writeXlsFile(data);
        return res.status(200).json({ message: 'Data updated successfully' });
      } else {
        return res.status(400).json({ error: 'Row index out of bounds' });
      }
    } else {
      return res.status(400).json({ error: 'Invalid input' });
    }
  }

  return res.status(404).json({ error: 'Not Found' });
};
