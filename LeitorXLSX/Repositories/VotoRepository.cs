using LeitorXLSX.Data;
using LeitorXLSX.Enums;
using LeitorXLSX.Interfaces;
using LeitorXLSX.Models;
using System.Runtime.InteropServices;
using static LeitorXLSX.Utils.Biblioteca;
using ExcelLeitorXLSX = Microsoft.Office.Interop.Excel;

namespace LeitorXLSX.Repositories
{
    public class VotoRepository : IVotoInterface
    {
        public readonly Context _context;

        public VotoRepository(Context context)
        {
            _context = context;
        }

        public async Task<IEnumerable<Voto>>? GetVotosSegundoTurno()
        {
            string caminhoXLSX = $"{AppContext.BaseDirectory}\\XLSX\\{GetDescricaoEnum(ListaXlsxEnum.SegundoTurno)}";

            IEnumerable<Voto> xlsxVotos = LerExcelSegundoTurno(caminho: caminhoXLSX, desconsiderarAteLinha: 2, isProcessoCompleto: false);

            if (xlsxVotos?.Count() > 0)
            {
                await _context.AddRangeAsync(xlsxVotos);
                await _context.SaveChangesAsync();
            }

            return xlsxVotos;
        }

        private static IEnumerable<Voto> LerExcelSegundoTurno(string caminho, int desconsiderarAteLinha, bool isProcessoCompleto)
        {
            // Tutorial de como "ler excel" em C#: https://coderwall.com/p/app3ya/read-excel-file-in-c
            List<Voto> xlsxVotos = new();

            // Criar referência do Excel;
            ExcelLeitorXLSX.Application xlApp = new();
            ExcelLeitorXLSX.Workbook xlWorkbook = xlApp.Workbooks.Open(caminho);
            ExcelLeitorXLSX._Worksheet xlWorksheet = (ExcelLeitorXLSX._Worksheet)xlWorkbook.Sheets[1];
            ExcelLeitorXLSX.Range xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            string nomeArquivo = Path.GetFileName(caminho);
            Console.WriteLine("\nForam encontradas " + rowCount + " linhas no arquivo " + nomeArquivo + "\nAguarde um momento");

            // Verificar se o processo deve ser completo ou não;
            int rows = isProcessoCompleto ? rowCount : 10; 
            var votos = GerarVotosYield(xlRange: xlRange, rows: rows, desconsiderarAteLinha: desconsiderarAteLinha);
          
            // Adicionar dados na lista final;
            xlsxVotos.AddRange(votos);

            // Etc;
            #region ETC
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            #endregion

            // Finalizar método retornando lista;
            return xlsxVotos;
        }

        private static IEnumerable<Voto> GerarVotosYield(ExcelLeitorXLSX.Range xlRange, int rows, int desconsiderarAteLinha)
        {
            // Loop em todas as linhas;
            for (int i = 1; i <= (rows + desconsiderarAteLinha); i++)
            {
                // Pular os cabeçalhos;
                if (i <= desconsiderarAteLinha)
                {
                    continue;
                }

                string UF = xlRange.Cells[i, 1].Value2 ?? "";
                string nomeMunicipio = xlRange.Cells[i, 2].Value2 ?? "";
                string qtdAptosMunicipio = xlRange.Cells[i, 3].Value2 ?? "";
                double codigoMunicipioIBGE = xlRange.Cells[i, 4].Value2 ?? 0;
                bool isCapital = xlRange.Cells[i, 5].Value2 == 1 ? true : false;
                double zona = xlRange.Cells[i, 6].Value2 ?? 0;
                double secao = xlRange.Cells[i, 7].Value2 ?? 0;
                double qtdAptos = xlRange.Cells[i, 8].Value2 ?? 0;
                double qtdVotos13 = xlRange.Cells[i, 9].Value2 ?? 0;
                double qtdVotos22 = xlRange.Cells[i, 10].Value2 ?? 0;
                double qtdTotalVotos1322 = xlRange.Cells[i, 11].Value2 ?? 0;
                double qtdVotosBranco = xlRange.Cells[i, 12].Value2 ?? 0;
                double qtdTotalFinal = xlRange.Cells[i, 13].Value2 ?? 0;

                Voto v = new()
                {
                    Turno = 2,
                    UF = UF,
                    NomeMunicipio = nomeMunicipio,
                    QtdAptosMunicipio = qtdAptosMunicipio,
                    CodigoMunicipioIBGE = codigoMunicipioIBGE,
                    IsCapital = isCapital,
                    Zona = zona,
                    Secao = secao,
                    QtdAptos = qtdAptos,
                    QtdVotos13 = qtdVotos13,
                    QtdVotos22 = qtdVotos22,
                    QtdTotalVotos1322 = qtdTotalVotos1322,
                    QtdVotosBranco = qtdVotosBranco,
                    QtdTotalFinal = qtdTotalFinal
                };

                Console.WriteLine($"Linha {(i - desconsiderarAteLinha)}/{rows} adicionada à lista final");

                yield return v;
            }
        }
    }
}
