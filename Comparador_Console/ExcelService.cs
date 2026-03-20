using Comparador_Console;
using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace Comparador_Console
{
    public static class ExcelService
    {
        static char DetectarSeparador(string linha)
        {
            if (linha.Contains(";")) return ';';
            if (linha.Contains("|")) return '|';
            if (linha.Contains(",")) return ',';

            return ';'; // padrão
        }

        public static List<RegistroGenerico> CarregarExcelGenerico(string caminho)
        {
            var lista = new List<RegistroGenerico>();
            var ext = Path.GetExtension(caminho)?.ToLowerInvariant();

            ExcelPackage.License.SetNonCommercialOrganization("ComparadorConsole");

            if (ext == ".csv")
            {
                var linhas = File.ReadAllLines(caminho);
                if (linhas.Length == 0) return lista;

                // detectar delimitador (',' ou ';')
                var delimiter = DetectarSeparador(linhas[0]);
                var headers = linhas[0].Split(delimiter);

                foreach (var linha in linhas.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(linha)) continue;
                    var cols = linha.Split(delimiter);
                    var reg = new RegistroGenerico();
                    for (int i = 0; i < headers.Length && i < cols.Length; i++)
                        reg.Campos[headers[i]] = cols[i];
                    lista.Add(reg);
                }
                return lista;
            }

            // tratar .xlsx e variantes suportadas pelo EPPlus
            using (var package = new ExcelPackage(new FileInfo(caminho)))
            {
                var ws = package.Workbook.Worksheets.FirstOrDefault();
                if (ws == null || ws.Dimension == null) return lista;

                int colunas = ws.Dimension.Columns;
                int linhas = ws.Dimension.Rows;

                var headers = new List<string>();
                for (int col = 1; col <= colunas; col++)
                    headers.Add(ws.Cells[1, col].Text);

                for (int lin = 2; lin <= linhas; lin++)
                {
                    var registro = new RegistroGenerico();
                    for (int col = 1; col <= colunas; col++)
                    {
                        var nomeColuna = headers[col - 1];
                        var valor = ws.Cells[lin, col].Text;
                        registro.Campos[nomeColuna] = valor;
                    }
                    lista.Add(registro);
                }
            }

            return lista;
        }

        public static List<RegistroGenerico> CarregarTxtGenerico(string caminho)
        {
            var lista = new List<RegistroGenerico>();

            var linhas = File.ReadAllLines(caminho);
            var separador = DetectarSeparador(linhas[0]);
            // cabeçalho
            var headers = linhas[0].Split(separador);

            for (int i = 1; i < linhas.Length; i++)
            {

                if (string.IsNullOrWhiteSpace(linhas[i]))
                    continue;

                var valores = linhas[i].Split(separador)
                    .Select(h => h.Trim())
                    .ToArray();

                var registro = new RegistroGenerico();

                if (valores.Length != headers.Length)
                {
                    Console.WriteLine($"Linha inválida: {linhas[i]}");
                    continue;
                }

                for (int j = 0; j < headers.Length; j++)
                {
                    var coluna = headers[j];
                    var valor = j < valores.Length ? valores[j] : "";

                    registro.Campos[coluna] = valor;
                }

                lista.Add(registro);
            }

            return lista;
        }

        public static List<RegistroGenerico> CarregarArquivoGenerico(string caminho)
        {
            var extensao = Path.GetExtension(caminho).ToLower();

            switch (extensao)
            {
                case ".xlsx":
                    return CarregarExcelGenerico(caminho);

                case ".csv":
                case ".txt":
                    return CarregarTxtGenerico(caminho);

                default:
                    throw new Exception("Formato de arquivo não suportado");
            }
        }
        //public static List<Registro> CarregarArquivoCsv(string caminho)
        //{
        //    var linhas = File.ReadAllLines(caminho);
        //    var lista = new List<Registro>();

        //    foreach (var linha in linhas.Skip(1))
        //    {
        //        var colunas = linha.Split(';'); // muda pra ',' se precisar

        //        var registro = new Registro
        //        {
        //            Id = colunas[0],
        //            GovernmentId = colunas[1],
        //            Cnpj = colunas[2]
        //        };

        //        lista.Add(registro);
        //    }

        //    return lista;
        //}

        //public static List<Registro> CarregarExcel(string caminho)
        //{
        //    var lista = new List<Registro>();

        //    ExcelPackage.License.SetNonCommercialOrganization("ComparadorConsole");

        //    using (var package = new ExcelPackage(new FileInfo(caminho)))
        //    {
        //        var worksheet = package.Workbook.Worksheets[0]; // primeira aba

        //        int linhas = worksheet.Dimension.Rows;

        //        for (int i = 2; i <= linhas; i++) // começa na linha 2 (pula cabeçalho)
        //        {
        //            if (string.IsNullOrEmpty(worksheet.Cells[i, 1].Text))
        //                continue;
        //            var registro = new Registro
        //            {
        //                Id = worksheet.Cells[i, 1].Text,
        //                GovernmentId = worksheet.Cells[i, 2].Text,
        //                Cnpj = worksheet.Cells[i, 3].Text
        //            };

        //            lista.Add(registro);
        //        }
        //    }

        //    return lista;
        //}

        public static void ExportarParaExcel(List<Diferenca> diferencas)
        {
            ExcelPackage.License.SetNonCommercialOrganization("ComparadorConsole");

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
                var basePath = Directory.GetCurrentDirectory();
                var projetoPath = Path.GetFullPath(Path.Combine(basePath, @"..\..\"));
                var caminhoResultado = Path.Combine(projetoPath, "Arquivos", "ResultadoComparacao.xlsx");
                var file = new FileInfo(caminhoResultado);
                package.SaveAs(file);
            }
        }

        public static string SelecionarArquivo(string titulo)
        {
            string result = string.Empty;
            var t = new Thread(() =>
            {

                using (var dialog = new OpenFileDialog())
                {
                    dialog.Title = titulo;
                    dialog.Filter = "Arquivos (*.xlsx;*.csv;*.txt)|*.xlsx;*.csv;*.txt|Todos (*.*)|*.*";

                    if (dialog.ShowDialog() == DialogResult.OK)
                    {
                        result = dialog.FileName;
                    }

                }
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
            return result;
        }
    }
}