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

        public async Task<List<Voto>>? GetVotosSegundoTurno()
        {
            string caminhoXLSX = $"{AppContext.BaseDirectory}\\XLSX\\{GetDescricaoEnum(ListaXlsxEnum.SegundoTurno)}";

            List<Voto> xlsxVotos = LerExcelSegundoTurno(caminhoXLSX);

            if (xlsxVotos?.Count > 0)
            {
                await _context.AddRangeAsync(xlsxVotos);
                // await _context.SaveChangesAsync();
            }

            return xlsxVotos;
        }

        private static List<Voto> LerExcelSegundoTurno(string caminho)
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

            // Loop em todas as linhas;
            for (int i = 1; i <= rowCount; i++)
            {
                // Pular o cabeçalho;
                if (i == 1)
                {
                    continue;
                }

                string teste = (string)(xlRange.Cells[i, 1]).Value2 ?? "";
                string teste2 = (string)(xlRange.Cells[i, 2]).Value2 ?? "";
                //string diretoria = (string)(xlRange.Cells[i, 3]).Value2;
                //string vereador = (string)(xlRange.Cells[i, 4]).Value2;
                //string tipoProtocolo = (string)(xlRange.Cells[i, 5]).Value2;
                //string assunto = (string)(xlRange.Cells[i, 6]).Value2;
                //string subdivisao = (string)(xlRange.Cells[i, 7]).Value2;
                //string regional = (string)(xlRange.Cells[i, 8]).Value2;
                //string numero = xlRange.Cells[i, 9].Value2.ToString(); // Double (?);
                //string complemento = (string)(xlRange.Cells[i, 10]).Value2;
                //string cep = (string)(xlRange.Cells[i, 11]).Value2;
                //string pontoReferencia = (string)(xlRange.Cells[i, 12]).Value2;
                //string descricao = (string)(xlRange.Cells[i, 13]).Value2;
                //string dadosImportantes = (string)(xlRange.Cells[i, 14]).Value2;
                //string status = (string)(xlRange.Cells[i, 15]).Value2;
                //string tipoDocExterno = (string)(xlRange.Cells[i, 16]).Value2;
                //string docExterno = (string)(xlRange.Cells[i, 17]).Value2;
                //string posicionamento = (string)(xlRange.Cells[i, 18]).Value2;
                //string dtResposta = xlRange.Cells[i, 19].Value2.ToString(); // Double (?)
                //string resposta = (string)(xlRange.Cells[i, 20]).Value2;
                //string statusRespostaParcial = (string)(xlRange.Cells[i, 21]).Value2;
                //string dtRespostaParcial = xlRange.Cells[i, 22].Value2.ToString(); // Double (?)
                //string respostaParcial = (string)(xlRange.Cells[i, 23]).Value2;
                //string sigiloso = (string)(xlRange.Cells[i, 24]).Value2;
                //string usuarioCadastro = (string)(xlRange.Cells[i, 25]).Value2;
                //string cpf = xlRange.Cells[i, 26].Value2.ToString();

                Voto v = new()
                {
                    Turno = 2,
                    NomeMunicipio = "",
                    QtdAptosMunicipio = "",
                    CodigoMunicipioIBGE = "",
                    IsCapital = false,
                    Zona = 0,
                    Secao = 0,
                    QtdAptos = 0,
                    QtdVotos13 = 0,
                    QtdVotos22 = 0,
                    QtdTotalVotos1322 = 0,
                    QtdVotosBranco = 0,
                    QtdTotalFinal = 0
                };

                xlsxVotos.Add(v);

                Console.WriteLine("Linha " + i + " adicionada na lista");
            }

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
    }
}
