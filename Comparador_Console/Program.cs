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
            Console.WriteLine("\n==============================");
            Console.WriteLine("📊 VISÃO GERAL DA COMPARAÇÃO");
            Console.WriteLine("==============================");

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


            string colunaId = "Customer_ID".ToUpper();
            var colunasComparar = new List<string>
            {
                "GovernmentId",
                "CRM_CNPJ"
            };
            var diferencas = new List<Diferenca>();

            if (!baseA.Any() || !baseA[0].Campos.ContainsKey(colunaId))
            {
                Console.WriteLine($"Coluna {colunaId} não encontrada no arquivo A");
                return;
            }

            var dictB = new Dictionary<string, RegistroGenerico>();

            foreach (var itemB in baseB)
            {
                if (!itemB.Campos.ContainsKey(colunaId))
                {
                    Console.WriteLine("[ERRO] Registro sem ID");
                    continue;
                }

                var id = itemB.Campos[colunaId];

                if (string.IsNullOrWhiteSpace(id))
                {
                    Console.WriteLine("[ERRO] ID vazio");
                    continue;
                }

                if (!dictB.ContainsKey(id))
                    dictB.Add(id, itemB);
                else
                    Console.WriteLine($"[DUPLICADO] ID: {id}");
            }

            foreach (var itemA in baseA)
            {
                if (!itemA.Campos.ContainsKey(colunaId))
                    continue;

                var id = itemA.Campos[colunaId];

                if (!dictB.ContainsKey(id))
                {
                    Console.WriteLine($"[SÓ NA A] ID: {id}");
                    continue;
                }

                var itemBComparar = dictB[id];

                foreach (var coluna in colunasComparar)
                {
                    var valorA = itemA.Campos.ContainsKey(coluna) ? itemA.Campos[coluna] : "";
                    var valorB = itemBComparar.Campos.ContainsKey(coluna) ? itemBComparar.Campos[coluna] : "";

                    if (!valorA.Equals(valorB, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine(
                            $"[DIVERGENTE] ID: {id} | Campo: {coluna} | A: {valorA} | B: {valorB}"
                        );

                        diferencas.Add(new Diferenca
                        {
                            Id = id,
                            Campo = coluna,
                            ValorA = valorA,
                            ValorB = valorB
                        });
          
                    }
                }
            }

            var divergentes = diferencas
                  .Select(x => x.Id)
                  .Distinct()
                  .Count();

            var idsA = baseA
                .Where(x => x.Campos.ContainsKey(colunaId) && !string.IsNullOrWhiteSpace(x.Campos[colunaId]))
                .Select(x => x.Campos[colunaId])
                .ToHashSet();

            var idsB = baseB
                .Where(x => x.Campos.ContainsKey(colunaId) && !string.IsNullOrWhiteSpace(x.Campos[colunaId]))
                .Select(x => x.Campos[colunaId])
                .ToHashSet();

            var apenasA = baseA
                .Where(a => a.Campos.ContainsKey(colunaId) && !idsB.Contains(a.Campos[colunaId]))
                .ToList();

            var apenasB = baseB
                .Where(b => b.Campos.ContainsKey(colunaId) && !idsA.Contains(b.Campos[colunaId]))
                .ToList();

            var soNaA = idsA.Except(idsB).Count();
            var soNaB = idsB.Except(idsA).Count();
            var emAmbas = idsA.Intersect(idsB).Count();
            var total = idsA.Union(idsB).Count();

            Console.WriteLine("\n📊 Visão geral");

            Console.WriteLine($"Registros apenas na Planilha A: {soNaA}");
            Console.WriteLine($"Registros apenas na Planilha B: {soNaB}");
            Console.WriteLine($"Registros presentes em ambas: {emAmbas}");
            Console.WriteLine($"Registros com alguma divergência (dados diferentes): {divergentes}");
            Console.WriteLine($"Divergência: {divergentes} ({(divergentes * 100.0 / total):F2}%)");

            ExcelService.ExportarParaExcel(diferencas, apenasA, apenasB);

            Console.WriteLine("\n✔ Comparação finalizada!");
            Console.WriteLine($"Total de divergências: {diferencas.Count}");

            Console.ReadLine();
        }
    }
}
