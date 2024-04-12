import express from 'express';
import bodyParser from 'body-parser';
import 'dotenv/config';
const app = express();

app.use(bodyParser.json());

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});