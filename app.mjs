import express from "express";
import { sendEmail, createOnlineMeeting } from './graphHelper.js';

const PORT = 8080;

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

app.post('/create-online-meeting', (req, res) => {
  const payload = req.body;
  createOnlineMeeting(payload);
  res.status(200).send(payload);
});

app.listen(PORT, () => console.log(`http://localhost:${PORT}`));