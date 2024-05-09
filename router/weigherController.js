const { weigherController } = require("../controllers");


const router = require("express").Router();

router.get("/weigher/export", weigherController.exportExcel);

module.exports = router;
