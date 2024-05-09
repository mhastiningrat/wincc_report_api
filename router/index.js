const summaryRouter = require("./summaryRouter");
const weigherRouter = require("./weigherController");
const lowrateRouter = require("./lowrateRouter");
const shippingRouter = require("./shippingRouter ");

module.exports = {
	route: [summaryRouter, weigherRouter, lowrateRouter, shippingRouter],
};
