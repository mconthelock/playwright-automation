// routes/index.js
const express = require("express");
const router = express.Router();

const meeting = require("./meeting/index");

router.use("/meeting", meeting);
module.exports = router;
