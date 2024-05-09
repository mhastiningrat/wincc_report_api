const { shippingController } = require("../controllers");

const router = require("express").Router();

router.get("/vessel/export", shippingController.exportExcel);

module.exports = router;
