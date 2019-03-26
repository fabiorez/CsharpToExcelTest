using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CsharpToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Gerando excel..");

            string posicaoColunaArea = "B";
            string posicaoUltimaColuna = "C";
            XLWorkbook wb = new XLWorkbook();
            IXLWorksheet ws = wb.Worksheets.Add("Planilha 1");
            //Cabeçalho do Relatório
            GerandoCabecalho(ws);

            //Corpo do relatório
            int linha = 2;
            string hierarquia = "CLARO > NIVEL 02";
            List<KeyValuePair<int, string>> listaHierarquia = new List<KeyValuePair<int, string>>();
            GerandoValores(ws, ref linha, ref hierarquia, listaHierarquia);

            //crio as colunas de areas
            if (listaHierarquia.Any())
            {
                posicaoUltimaColuna = CriandoColunaAreas(posicaoColunaArea, ws, listaHierarquia);
            }

            //ajuste da linha
            linha--;

            //crio tabela para ativar os filtros
            IXLRange range = ws.Range($"A1:{posicaoUltimaColuna}{linha}");
            range.CreateTable();

            //ajuste de tamanho apos preencher as colunas
            int intPosicaoUltimaColuna = StringToInt(posicaoUltimaColuna);
            ws.Columns($"1-{intPosicaoUltimaColuna}").AdjustToContents();

            //remove coluna area
            ws.Column(StringToInt(posicaoColunaArea)).Delete();

            //Salvar arquivo
            wb.SaveAs(@"c:\Projetos Git\teste.xlsx");

            //Libera a memoria
            wb.Dispose();

            Console.WriteLine("Feito!");
            Console.ReadKey();
        }

        private static string CriandoColunaAreas(string posicaoColunaArea, 
                                                 IXLWorksheet ws, 
                                                 List<KeyValuePair<int, string>> listaHierarquia)
        {
            string posicaoUltimaColuna;
            List<string> areas = new List<string>();
            List<string> nomeColunasInseridas = new List<string>();
            int colunaInserir = StringToInt(posicaoColunaArea);

            foreach (KeyValuePair<int, string> item in listaHierarquia)
            {
                int linhaCelula = item.Key;
                areas = item.Value.Split('>').Select(x => x.Trim()).ToList();
                int qtdeAreas = areas.Count();
                int colunaCelula = StringToInt(posicaoColunaArea) + 1;
                GerandoColunas(ws, nomeColunasInseridas, colunaInserir, qtdeAreas);
                InserindoDadosColunas(ws, areas, linhaCelula, qtdeAreas, colunaCelula);
            }

            posicaoUltimaColuna = NumberToString(colunaInserir + 1);
            return posicaoUltimaColuna;
        }

        private static void GerandoCabecalho(IXLWorksheet ws)
        {
            ws.Cell("A1").Value = "Item";
            ws.Cell("B1").Value = "Área";
            ws.Cell("C1").Value = "Nome";
        }

        private static void GerandoColunas(IXLWorksheet ws, 
                                           List<string> nomeColunasInseridas, 
                                           int colunaInserir, 
                                           int qtdeAreas)
        {
            for (int posicaoArea = 1; posicaoArea < qtdeAreas; posicaoArea++)
            {
                string tituloColuna = $"Área Nível {posicaoArea}";
                bool naoExiste = !nomeColunasInseridas.Any(x => x == tituloColuna);

                if (naoExiste)
                {
                    int linhaTitulo = 1;
                    int quantidadeColunasInserir = 1;
                    nomeColunasInseridas.Add(tituloColuna);
                    ws.Column(colunaInserir).InsertColumnsAfter(quantidadeColunasInserir);
                    IXLCell titulo = ws.Cell(linhaTitulo, colunaInserir + 1);
                    titulo.Value = tituloColuna;
                    colunaInserir++;
                }
            }
        }

        private static void GerandoValores(IXLWorksheet ws, 
                                           ref int linha, 
                                           ref string hierarquia, 
                                           List<KeyValuePair<int, string>> listaHierarquia)
        {
            for (int i = 1; i <= 8; i++)
            {
                ws.Cell($"A{linha}").Value = $"{i}";

                if (i != 1)
                {
                    hierarquia = $"{hierarquia} > NIVEL 0{i + 1}";
                }

                ws.Cell($"B{linha}").Value = hierarquia;

                ws.Cell($"C{linha}").Value = $"João Testador {i}";

                listaHierarquia.Add(new KeyValuePair<int, string>(i, hierarquia));

                linha++;
            }
        }

        private static void InserindoDadosColunas(IXLWorksheet ws, 
                                                  List<string> areas, 
                                                  int linhaCelula, 
                                                  int qtdeAreas, 
                                                  int colunaCelula)
        {
            for (int posicaoArea = 1; posicaoArea < qtdeAreas; posicaoArea++)
            {
                IXLCell celulaInserir = ws.Cell(linhaCelula + 1, colunaCelula);
                int posicaoAreaInserir = posicaoArea - 1;
                celulaInserir.Value = areas[posicaoAreaInserir].Trim();
                colunaCelula++;
            }
        }
                     
        private static string NumberToString(int colunaInserir)
        {
            Char coluna = (Char)((colunaInserir + 64));

            return coluna.ToString();
        }

        private static int StringToInt(string posicaoUltimaColuna)
        {
            char[] c = posicaoUltimaColuna.ToCharArray();
            return char.ToUpper(c[0]) - 64;
        }
    }
}
