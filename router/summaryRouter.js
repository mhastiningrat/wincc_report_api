const { summaryController } = require("../controllers");

const router = require("express").Router();

router.get("/summary-page1/export", summaryController.exportExcel);
router.get("/summary-page2/export", summaryController.exportExcelPage2);

module.exports = router;
