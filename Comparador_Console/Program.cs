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
        static void Main(string[] args)
        {
            var baseA = CarregarArquivo("C:\\Users\\MSilva\\Documents\\Estudos\\Documentos de estudo\\Projeto\\Comparador_Console\\Comparador_Console\\Comparador_Console\\Arquivos\\A.csv");
            var baseB = CarregarArquivo("C:\\Users\\MSilva\\Documents\\Estudos\\Documentos de estudo\\Projeto\\Comparador_Console\\Comparador_Console\\Comparador_Console\\Arquivos\\B.csv");
            var diferencas = new List<Diferenca>();

            var dictB = baseB.ToDictionary(x => x.Id);

            foreach (var itemA in baseA)
            {
                if (!dictB.ContainsKey(itemA.Id))
                {
                    Console.WriteLine($"Só existe na base A: {itemA.Id}");
                    continue;
                }

                var itemB = dictB[itemA.Id];

                if (itemA.GovernmentId != itemB.GovernmentId)
                {
                    Console.WriteLine(
                        $" - GovernmentId | A: {itemA.GovernmentId} | B: {itemB.GovernmentId}"
                    );
                    diferencas.Add(new Diferenca
                    {
                        Id=itemA.Id,
                        Campo = "GovernmentId",
                        ValorA= itemA.GovernmentId,
                        ValorB = itemB.GovernmentId,
                    });
                }

                if (itemA.Cnpj != itemB.Cnpj)
                {
                    if (itemA.Cnpj != itemB.Cnpj)
                    {
                        Console.WriteLine(
                            $" - CNPJ         | A: {itemA.Cnpj} | B: {itemB.Cnpj}"
                        );
                    }
                    diferencas.Add(new Diferenca
                    {
                        Id = itemA.Cnpj,
                        Campo = "CNPJ",
                        ValorA = itemA.Cnpj,
                        ValorB = itemB.Cnpj,
                    });
                }
            }

            var idsA = baseA.Select(x => x.Id);
            var idsB = baseB.Select(x => x.Id);

            var soNaB = idsB.Except(idsA);

            foreach (var id in soNaB)
            {
                Console.WriteLine($"Só existe na base B: {id}");
            }

            ExcelService.ExportarParaExcel(diferencas);

            Console.WriteLine("\n✔ Comparação finalizada!");
            Console.WriteLine($"Total de divergências: {diferencas.Count}");

            Console.ReadLine();
        }
        static List<Registro> CarregarArquivo(string caminho)
        {
            var linhas = File.ReadAllLines(caminho);
            var lista = new List<Registro>();

            foreach (var linha in linhas.Skip(1))
            {
                var colunas = linha.Split(';'); // muda pra ',' se precisar

                var registro = new Registro
                {
                    Id = colunas[0],
                    GovernmentId = colunas[1],
                    Cnpj = colunas[2]
                };

                lista.Add(registro);
            }

            return lista;
        }
    }
}
