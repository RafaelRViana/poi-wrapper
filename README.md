Bem-vindo a biblioteca poi-wrapper
============================

Este projeto foi criado com o objetivo de facilitar a utilização da biblioteca Apache POI.

## Exemplos ##

### Leitura de um Arquivo Excel ###

String arquivo = "localização do arquivo";

//Excel File representa a estrutura de uma planilha do Excel. Ele é carregado através de um File.
ExcelFile arquivoExcel = new ExcelFile(new FileInputStream(arquivo));

//Excel Sheet representa uma planilha do Excel, é obtida através do ExcelFile
ExcelSheet planilha = arquivoExcel.getSheet(0);

//Excel Extrator é uma classe que auxilia a extração de informações das planilhas de Excel, com tratamentos de erros que podem ser causados pela biblioteca Apache POI
ExcelExtrator extrator = new ExcelExtrator(planilha);

//Para fazer a leitura das 1000 linhas do arquivo, vamos utilizar
for(int linha = 1; linha <= 1000; linha++) {
  extrator.getStringOrNumberAsString("A"+linha); //O método recebe o parametro de uma string que é equivalente a forma legível de uma coluna no Excel A1, B1, C1, ...
}
