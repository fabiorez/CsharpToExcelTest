using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace CsharpToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Gerando excel..");

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Planilha 1");
            string posicaoUltimaColuna = "B";

            //Cabeçalho do Relatório
            ws.Cell("A1").Value = "Item";
            ws.Cell("B1").Value = "Área";

            //Corpo do relatório
            int linha = 2;

            string hierarquia = "CLARO > NIVEL 01";
            List<string> hierarquias = new List<string>();

            for (int i = 1; i <= 8; i++)
            {
                ws.Cell($"A{linha}").Value = $"{i}";

                if (i != 1)
                {
                    hierarquia = $"{hierarquia} > NIVEL 0{i}";
                }

                ws.Cell($"B{linha}").Value = hierarquia;

                //aqui separo a hierarquia
                hierarquias = hierarquia.Split('>').Select(x => x.Trim()).ToList();

                int quantidadeHierarquias = hierarquias.Count();

                for (int j = 0; j < quantidadeHierarquias; j++)
                {
                    ws.Range($"A1:{posicaoUltimaColuna}{linha}").InsertColumnsAfter(1);
                    posicaoUltimaColuna++;
                }

                linha++;
            }

            //ajuste da linha
            linha--;

            //crio tabela para ativar os filtros
            IXLRange range = ws.Range($"A1:B{linha}");
            range.CreateTable();

            //ajuste de tamanho apos preencher as colunas
            ws.Columns("1-2").AdjustToContents();

            //Salvar arquivo
            wb.SaveAs(@"c:\Projetos Git\teste.xlsx");

            //Libera a memoria
            wb.Dispose();

            Console.WriteLine("Feito!");
            Console.ReadKey();
        }
    }
}
