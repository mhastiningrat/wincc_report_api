const express = require("express");
const bodyParser = require("body-parser");
const dotenv = require("dotenv");
const app = express();
const router = require("./router");
dotenv.config();
app.use(bodyParser.json());

for (route of router.route) {
	app.use("/api", route);
}

const port = process.env.PORT || 1234;
app.listen(port, () => {
	console.log(`Server is running on port ${port}`);
});
