import express from 'express';
import handler from './api/submit.js';

const app = express();
app.use(express.json());
app.post('/api/submit', handler);

app.listen(8080, () => {
  console.log('Server running on port 8080');
});
