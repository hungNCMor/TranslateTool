using IronXL;
using Microsoft.Extensions.Logging;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TranslateLib.Interface;
using TranslateLib.Model;

namespace TranslateLib
{
    public class TranslateExcelWithIronXL : ITranslateExcel
    {
        ILogger _logger;
        ITranslate _translate;
        public TranslateExcelWithIronXL(ILogger logger, ITranslate translate)
        {
            _logger = logger;
            _translate = translate;

        }
        public async Task<MemoryStream> TranslateExcelByStream(MemoryStream stream, string fileName)
        {


            WorkBook workBook = WorkBook.Load(stream);
            // Loop through each sheet in the workbook
            // Get the worksheets
            List<WorkSheet> worksheets = workBook.WorkSheets.ToList();
            int i = 0;
            var tasks = new List<Task>();
            // Iterate through the worksheets
            foreach (WorkSheet sheet in worksheets)
            {
                tasks.Add(HandleWorkSheet(sheet, i));
                i++;
            }
            var t1 = DateTime.Now;
            await Task.WhenAll(tasks);
            _logger.LogInformation("GetValue time " + (DateTime.Now - t1));
            var fileName1 = _translate.TranslateText(fileName);
            var t2 = DateTime.Now;
            //UpdateWorkBook(workBook, tasks);
            _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
            return workBook.ToStream();


        }
        private async Task HandleWorkSheet(WorkSheet sheet, int i)
        {
            try
            {
                string sheetName = _translate.TranslateText(sheet.Name).Replace("/", ".").Replace("・", ".");

                sheetName = sheetName.Replace('[', '(').Replace(']', ')').Substring(0, sheetName.Length > 29 ? 30 : sheetName.Length);
                sheet.Name = sheetName + i;
                var result = new List<DataResult>();

                // Get the last row and column indexes in the sheet
                int lastRow = sheet.Rows.Count();
                int lastColumn = sheet.Columns.Count();

                // Iterate through each row
                for (int row = 0; row < lastRow; row++)
                {
                    // Iterate through each column
                    for (int column = 0; column < lastColumn; column++)
                    {
                        string cellValue;
                        // Read the value of the cell
                        if (sheet.GetCellAt(row, column) != null)
                        {
                            if (!sheet.GetCellAt(row, column).IsFormula)
                            {
                                _logger.LogInformation($"GetCellAt row={row}, column={column}, sheet ={sheet.Name}");
                                cellValue = sheet.GetCellAt(row, column)?.StringValue;

                                if (String.IsNullOrEmpty(cellValue?.Trim()))
                                {
                                    continue;
                                }
                                _logger.LogInformation($"cellValue={cellValue}");
                                sheet.GetCellAt(row, column).Value = _translate.TranslateText(cellValue).Replace("/", ".");
                            }
                        }
                    }
                }
                i++;
                sheet.UnprotectSheet();
            }
            catch (Exception x)
            {
                throw x;
            }
        }
        public async Task<MemoryStream> TranslateExcelByPath(string path)
        {
            try
            {
                var t = DateTime.Now;
                var fileName = path.Replace(".xlsx", "").Split('\\').Last();
                Console.WriteLine($" {fileName} ThreadId: {System.Environment.CurrentManagedThreadId}");

                using (var ms = new MemoryStream())
                {
                    WorkBook workBook = WorkBook.Load(path);

                    // Loop through each sheet in the workbook
                    // Get the worksheets
                    List<WorkSheet> worksheets = workBook.WorkSheets.ToList();
                    int i = 0;
                    var tasks = new List<Task>();
                    // Iterate through the worksheets
                    foreach (WorkSheet sheet in worksheets)
                    {
                        tasks.Add(HandleWorkSheet(sheet, i));
                        i++;
                    }
                    var t1 = DateTime.Now;
                    await Task.WhenAll(tasks);
                    _logger.LogInformation("GetValue time " + (DateTime.Now - t1));
                    var t2 = DateTime.Now;
                    //UpdateWorkBook(workBook, tasks);
                    _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
                    _logger.LogInformation("total time " + (DateTime.Now - t));
                    return workBook.ToStream();
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public async Task TranslateExcelByPathSavePath(string path)
        {
            try
            {
                var t = DateTime.Now;
                var fileName = path.Replace(".xlsx", "").Split('\\').Last();
                Console.WriteLine($" {fileName} ThreadId: {System.Environment.CurrentManagedThreadId}");

                using (var ms = new MemoryStream())
                {
                    WorkBook workBook = WorkBook.Load(path);

                    // Loop through each sheet in the workbook
                    // Get the worksheets
                    List<WorkSheet> worksheets = workBook.WorkSheets.ToList();
                    int i = 0;
                    var tasks = new List<Task>();
                    // Iterate through the worksheets
                    foreach (WorkSheet sheet in worksheets)
                    {
                        tasks.Add(HandleWorkSheet(sheet, i));
                        i++;
                    }
                    var t1 = DateTime.Now;
                    await Task.WhenAll(tasks);
                    _logger.LogInformation("GetValue time " + (DateTime.Now - t1));
                    var t2 = DateTime.Now;
                    //UpdateWorkBook(workBook, tasks);
                    _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
                    var newFileName = fileName + _translate.TranslateText(fileName) + ".xlsx";
                    workBook.SaveAs(Path.Combine(string.Join("\\", path.Replace(".xlsx", "").Split('\\').SkipLast(1)), newFileName));
                }
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
