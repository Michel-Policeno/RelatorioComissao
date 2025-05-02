const express = require("express");
const router = express.Router();
const multer = require("multer");
const upload = require("../config/multer.config.js");
const relatorioController = require("../controller/relatorio.controller.js");

router.get("/", relatorioController.validaConexao);
router.post("/upload", upload.single("file"), relatorioController.uploadFile);

module.exports = router;
