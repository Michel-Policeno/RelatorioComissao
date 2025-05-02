const xlsx = require("xlsx");
const fs = require("fs");
const path = require("path");

const relatorioController = {
  // GET /
  validaConexao: (req, res) => {
    const objeto = { conexão: "concluida", data: new Date() };
    return res.json(objeto);
  },

  // POST /upload
  uploadFile: async (req, res) => {
    if (!req.file) {
      return res.status(400).json({ mensagem: "arquivo excel é obrigatorio" });
    }

    const caminhoArquivoXLSX = path.join(
      __dirname,
      "../upload",
      req.file.filename
    );

    const regexSepararExtensao = /\.[a-zA-Z0-9]+$/gm;
    const extensaoArquivo = regexSepararExtensao.exec(req.file.originalname);
    if (
      extensaoArquivo[0] !== ".xlsx" &&
      extensaoArquivo[0] !== ".xls" &&
      extensaoArquivo[0] !== ".xlsm"
    ) {
      fs.unlinkSync(caminhoArquivoXLSX); //excluir arquivo em formato invalido
      return res.status(400).json({
        mensagem: `Formato do arquivo inválido! Extensões permitidas: .xlsx .xls .xlsm`,
      });
    }
    //leitura do arquivo - bibliotexa xlsx
    try {
      const workbook = xlsx.readFile(caminhoArquivoXLSX);
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const data = xlsx.utils.sheet_to_json(sheet);

      const agrupados = data.reduce((acumulador, linhaExcel) => {
        const gerente = linhaExcel.Produtor;
        if (!acumulador[gerente]) acumulador[gerente] = [];
        acumulador[gerente].push(linhaExcel);
        return acumulador;
      }, {});
      res.json({ dados: agrupados });
    } catch (error) {
      return res.status(400).json({
        mensagem: `Erro no processamento do arquivo. ERRO: ${error.message}`,
      });
    } finally {
      fs.unlinkSync(caminhoArquivoXLSX);
    }
  },
};

module.exports = relatorioController;
