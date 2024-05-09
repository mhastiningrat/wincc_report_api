const { lowrateController } = require("../controllers");

const router = require("express").Router();

router.get("/lowrate/export", lowrateController.exportExcel);

module.exports = router;
