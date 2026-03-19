using Comparador_Console;
using OfficeOpenXml;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace Comparador_Console
{
    public static class ExcelService
    {
        public static void ExportarParaExcel(List<Diferenca> diferencas)
        {
            ExcelPackage.License = new OfficeOpenXml.License.NonCommercialLicense();
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Diferencas");

                worksheet.Cells[1, 1].Value = "ID";
                worksheet.Cells[1, 2].Value = "Campo";
                worksheet.Cells[1, 3].Value = "Valor A";
                worksheet.Cells[1, 4].Value = "Valor B";

                int linha = 2;

                foreach (var d in diferencas)
                {
                    worksheet.Cells[linha, 1].Value = d.Id;
                    worksheet.Cells[linha, 2].Value = d.Campo;
                    worksheet.Cells[linha, 3].Value = d.ValorA;
                    worksheet.Cells[linha, 4].Value = d.ValorB;
                    linha++;
                }

                var file = new FileInfo("ResultadoComparacao.xlsx");
                package.SaveAs(file);
            }
        }
    }
}