const express = require('express');
const app = express();

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'https://rpgnet.sharepoint.com');
  next();
});

app.get('/api', (req, res) => {
  res.json({ message: 'Hello from server!' });
});

app.listen(3000, () => console.log('Server started'));