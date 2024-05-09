const { summaryController } = require("../controllers");

const router = require("express").Router();

router.get("/summary/export", summaryController.exportExcel);

module.exports = router;
