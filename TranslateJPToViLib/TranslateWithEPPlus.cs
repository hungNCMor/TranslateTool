using IronXL;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TranslateJPToViLib;
using TranslateLib.Interface;
using TranslateLib.Model;
using OfficeOpenXml;

namespace TranslateLib
{
    public class TranslateWithEPplus : ITranslateExcel
    {
        ILogger _logger;
        ITranslate _translate;
        public TranslateWithEPplus(ILogger logger, ITranslate translate)
        {
            _logger = logger;
            _translate = translate;

        }

        public Task<MemoryStream> TranslateExcelByPath(string path)
        {
            // Assuming you have an Excel package object called 'package'
            ExcelPackage package = new ExcelPackage(new FileInfo("path/to/your/file.xlsx"));

            // Assuming you have a worksheet object called 'worksheet'
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];

            // Read the value from a specific cell
            int rowNumber = 1; // Example row number
            int columnNumber = 1; // Example column number
            object cellValue = worksheet.Cells[rowNumber, columnNumber].Value;

            // Close the package
            package.Dispose();
            throw new NotImplementedException();
        }

        public Task TranslateExcelByPathSavePath(string path)
        {
            throw new NotImplementedException();
        }

        public Task<MemoryStream> TranslateExcelByStream(MemoryStream stream, string fileName)
        {
            throw new NotImplementedException();
        }
    }
}
