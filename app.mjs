import express from "express";
import { sendEmail } from './sendMail.js';

const app = express();

app.use(express.json());

app.use(express.static('public'));

app.get('/', (req, res) => {
  res.status(200).send(`test`);
});

app.post('/create-event', (req, res) => {
  const payload = req.body;
  sendEmail(payload);
  res.status(200).send(payload);
});

const PORT = 80;

app.listen(PORT, () => console.log(`http://localhost:${PORT}`));