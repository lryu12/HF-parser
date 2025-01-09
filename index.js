const express = require('express');
const bodyParser = require('body-parser');
const parseSpreadsheet = require('./utils/ssParser');

const app = express();
const PORT = 3000;

app.use(bodyParser.json());

app.post('/parse', async (req, res) => {
  const filePath = './data/sample.xlsx'; // Example file
  try {
    const result = await parseSpreadsheet(filePath);
    res.json(result);
  } catch (err) {
    res.status(500).send({ error: 'Failed to parse spreadsheet' });
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
