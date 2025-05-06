const express = require("express");
const relatorioRoute = require("./router/relatorio.route.js");
const cors = require("cors");
const bodyParser = require("body-parser");
const path = require("node:path");

const app = express();
app.use(cors());
app.use(bodyParser.text({ type: "text/html", limit: "10mb" }));
app.use("./relatorio-static", express.static(path.resolve("dist")));
console.log(path.resolve("dist"));
app.use(express.json()); //habilita json
app.use(express.urlencoded({ extended: true })); //configuração ler dados da requisição

//rotas
app.use(relatorioRoute);

const PORT = 3000;
app.listen(PORT, () =>
  console.log(`Servidor Iniciado com sucesso. http://localhost:${PORT}/`)
);
