
using Microsoft.Extensions.Logging;

using TranslateLib.Interface;
using FastExcel;
using DocumentFormat.OpenXml.Spreadsheet;

namespace TranslateLib
{
    public class TranslateWithFastExcel :ITranslateExcel
    {
        ILogger _logger;
        ITranslate _translate;
        public TranslateWithFastExcel(ILogger logger, ITranslate translate)
        {
            _logger = logger;
            _translate = translate;

        }

        public Task<MemoryStream> TranslateExcelByPath(string name)
        {
            throw new NotImplementedException();
        }

        public async Task TranslateExcelByPathSavePath(string path)
        {// Get the input file path
            var inputFile = new FileInfo(path);
            // Create an instance of Fast Excel
            using (FastExcel.FastExcel fastExcel = new FastExcel.FastExcel(inputFile, true))
            { 
                foreach (var worksheet in fastExcel.Worksheets)
                {
                    Console.WriteLine(string.Format("Worksheet Name:{0}, Index:{1}", worksheet.Name, worksheet.Index));

                    //To read the rows call read
                    worksheet.Read();
                    var rows = worksheet.Rows.ToArray();
                    //Do something with rows
                    foreach (var row in rows)
                    {
                        //foreach (Cell cell in row.Cells)
                        //{
                        //    string cellValue = worksheet.GetCellValue(sheetName, rowNumber, columnNumber);
                        //    cell.Value.ToString();
                        //}
                    }
                    Console.WriteLine(string.Format("Worksheet Rows:{0}", rows.Count()));
                }
            }
        }

        public Task<MemoryStream> TranslateExcelByStream(MemoryStream stream, string fileName)
        {
            throw new NotImplementedException();
        }
    }
}
