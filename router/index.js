const summaryRouter = require("./summaryRouter");
const weigherRouter = require("./weigherController")

module.exports = {
	route: [summaryRouter,weigherRouter],
};
