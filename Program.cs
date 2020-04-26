using System;
using NPOI.HSSF.UserModel;
using System.IO;
using System.Collections.Generic;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace GerarExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            GerarExcel(TipoExcel.xlsx);
            GerarExcel(TipoExcel.xls);
            GerarExcel(TipoExcel.csv);
        }

        public static void GerarExcel(TipoExcel tipoExcel)
        {
            // Conteudo
            var items = new List<Pessoa> { new Pessoa { Id = "1", Nome = "Rafael jose" }, new Pessoa { Id = "2", Nome = "Thayani da Silva" } };

            //Tipo de Arquivo Excel
            var workbook = (IWorkbook)null;

            //planilha dentro do excel
            var sheet = (ISheet)null;

            // Definir nome do cabeçalho que sera definido na primeira linha do Excel
            var headers = new[] { "Id", "Nome" };

            // Definir nome da coluna que esse nome da coluna usa para buscar dados da lista
            var columns = new[] { "Id", "Nome" };

            //diretorio onde sera criado a planilha
            string filePath = "teste_" + DateTime.Now.ToString("yyyyMMddHHmmss") + "." + tipoExcel.ToString();

            if (tipoExcel == TipoExcel.xlsx)
            {
                // Declara o objeto XSSFWorkbook para criar o sheet
                workbook = new XSSFWorkbook();
            }
            else
            {
                // Declara o objeto HSSFWorkbook para criar o sheet
                workbook = new HSSFWorkbook();
            }

            sheet = workbook.CreateSheet("NomeMinhaPlanilha");

            var headerRow = sheet.CreateRow(0);

            // O loop abaixo é criar cabecalho
            for (int i = 0; i < columns.Length; i++)
            {
                var cell = headerRow.CreateCell(i);
                cell.SetCellValue(headers[i]);
            }

            // O loop abaixo é o conteúdo de preenchimento
            for (int i = 0; i < items.Count; i++)
            {
                var rowIndex = i + 1;
                var row = sheet.CreateRow(rowIndex);

                for (int j = 0; j < columns.Length; j++)
                {
                    var cell = row.CreateCell(j);
                    var o = items[i];
                    cell.SetCellValue(o.GetType().GetProperty(columns[j]).GetValue(o, null).ToString());
                }
            }

            // Criar o Arquivo
            FileStream file = File.Create(filePath);
            workbook.Write(file);
            file.Close();
        }

    }

    public class Pessoa
    {
        public string Id { get; set; }
        public string Nome { get; set; }
    }

    public enum TipoExcel
    {
        csv,
        xls,
        xlsx
    }

}
