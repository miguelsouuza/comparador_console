using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Comparador_Console
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("Selecione o arquivo A:");
            var caminhoA = ExcelService.SelecionarArquivo("Selecione o arquivo da Base A");

            Console.WriteLine("Selecione o arquivo B:");
            var caminhoB = ExcelService.SelecionarArquivo("Selecione o arquivo da Base B");

            if (string.IsNullOrEmpty(caminhoA) || string.IsNullOrEmpty(caminhoB))
            {
                Console.WriteLine("Arquivo não selecionado.");
                return;
            }


            var baseA = ExcelService.CarregarArquivoGenerico(caminhoA);
            var baseB = ExcelService.CarregarArquivoGenerico(caminhoB);

            string colunaId = "Id";

            var colunasComparar = new List<string>
            {
                "GovernmentId",
                "Cnpj"
            };

            var diferencas = new List<Diferenca>();

            var dictB = baseB.ToDictionary(x => x.Campos[colunaId]);

            foreach (var itemA in baseA)
            {
                var id = itemA.Campos[colunaId];

                if (!dictB.ContainsKey(id))
                {
                    Console.WriteLine($"[SÓ NA A] ID: {id}");
                    continue;
                }

                var itemB = dictB[id];

                foreach (var coluna in colunasComparar)
                {
                    var valorA = itemA.Campos.ContainsKey(coluna) ? itemA.Campos[coluna] : "";
                    var valorB = itemB.Campos.ContainsKey(coluna) ? itemB.Campos[coluna] : "";

                    if (valorA != valorB)
                    {
                        Console.WriteLine(
                            $"[DIVERGENTE] ID: {id} | Campo: {coluna} | A: {valorA} | B: {valorB}"
                        );
                    }
                    diferencas.Add(new Diferenca
                    {
                        Id = id,
                        Campo = coluna,
                        ValorA = valorA,
                        ValorB = valorB
                    });
                }
            }
            var idsA = baseA.Select(x => x.Campos["Id"]);
            var idsB = baseB.Select(x => x.Campos["Id"]);

            var soNaB = idsB.Except(idsA);

            foreach (var id in soNaB)
            {
                Console.WriteLine($"[SÓ NA B] ID: {id}");
            }

            ExcelService.ExportarParaExcel(diferencas);

            Console.WriteLine("\n✔ Comparação finalizada!");
            Console.WriteLine($"Total de divergências: {diferencas.Count}");

            Console.ReadLine();
        }
    }
}
