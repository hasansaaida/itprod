const express = require('express');
const bodyParser = require('body-parser');
const sql = require('mssql');

const app = express();
app.use(bodyParser.json());

// הגדרות חיבור ל-SQL Server
const config = {
  user: 'your_username',
  password: 'your_password',
  server: 'localhost', // או כתובת השרת שלך
  database: 'your_database',
  options: {
    trustServerCertificate: true,
    encrypt: false
  }
};

// נקודת API לקליטת ציוד חדש
app.post('/api/assets', async (req, res) => {
  const { name, vendor, warranty_expiry, status, barcode, history } = req.body;

  try {
    await sql.connect(config);
    const result = await sql.query`
      INSERT INTO Assets (name, vendor, warranty_expiry, status, barcode, history)
      VALUES (${name}, ${vendor}, ${warranty_expiry}, ${status}, ${barcode}, ${history})
    `;
    res.status(200).send({ message: 'הציוד נשמר בהצלחה!' });
  } catch (err) {
    console.error('שגיאה:', err);
    res.status(500).send({ error: 'שגיאה בשמירת הציוד' });
  }
});

// הפעלת השרת
app.listen(5000, () => {
  console.log('השרת רץ על פורט 5000');
});
