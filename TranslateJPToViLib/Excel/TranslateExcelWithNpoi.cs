using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Microsoft.Extensions.Logging;
using TranslateLib.Interface;
using TranslateLib.Model;
using DocumentFormat.OpenXml.Office2021.DocumentTasks;
using Task = System.Threading.Tasks.Task;
using NPOI.HPSF;
using DocumentFormat.OpenXml.Spreadsheet;
using CellType = NPOI.SS.UserModel.CellType;

namespace TranslateLib.Excel
{
    public class TranslateExcelWithNpoi : ITranslateExcel
    {
        ILogger<TranslateExcelWithNpoi> _logger;
        ITranslate _translate;
        public TranslateExcelWithNpoi(ILogger<TranslateExcelWithNpoi> logger, ITranslate translate)
        {
            _logger = logger;
            _translate = translate;

        }

        public Task<MemoryStream> TranslateExcelByPath(string path)
        {
            throw new NotImplementedException();
        }
        private async Task HandleWorkSheet(ISheet sheet, int i)
        {
            try
            {
                // Iterate through the rows in the sheet
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    if (row != null)
                    {
                        // Iterate through the cells in the row
                        for (int columnIndex = 0; columnIndex < row.LastCellNum; columnIndex++)
                        {
                            ICell cell = row.GetCell(columnIndex);

                            if (cell != null && (cell.CellType == CellType.Unknown || cell.CellType == CellType.String))
                            {
                                // Retrieve the cell value
                                string cellValue = cell.ToString().Trim();
                                if (!string.IsNullOrEmpty(cellValue))
                                    cell.SetCellValue(_translate.TranslateText(cellValue!));
                                Console.WriteLine($"Cell value: {cellValue}");
                            }
                        }
                    }
                }
            }
            catch (Exception x)
            {
                throw x;
            }
        }
        public async Task TranslateExcelByPathSavePath(string path)
        {
            try
            {
                IWorkbook workbook;
                var fileName = path.Replace(".xlsx", "").Split('\\').Last();
                using (var fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    workbook = new XSSFWorkbook(fileStream); // For .xlsx files
                                                             // workbook = new HSSFWorkbook(fileStream); // For .xls files
                    int sheetCount = workbook.NumberOfSheets;
                    var tasks = new List<Task>();
                    for (int i = 0; i < sheetCount; i++)
                    {
                        ISheet sheet = workbook.GetSheetAt(i);
                        string sheetName = sheet.SheetName;
                        string newSheetName = _translate.TranslateText(sheet.SheetName, "ja", "en").Replace("/", ".").Replace("・", ".");

                        newSheetName = newSheetName.Replace('[', '(').Replace(']', ')').Substring(0, newSheetName.Length > 29 ? 29 : newSheetName.Length) + i;
                        //newSheetName = sheetName.Length <= 29 ? newSheetName : newSheetName + i;
                        tasks.Add(HandleWorkSheet(sheet, i));
                        workbook.SetSheetName(i, newSheetName);
                        var t1 = DateTime.Now;
                        await Task.WhenAll(tasks);
                    }
                    var newFileName = fileName + _translate.TranslateText(fileName,"ja","en").Replace("/", "_").Replace("\\", "_") + ".xlsx";
                    using (FileStream fs = new FileStream(Path.Combine(string.Join("\\", path.Replace(".xlsx", "").Split('\\').SkipLast(1)), newFileName), FileMode.CreateNew))
                        workbook.Write(fs);
                }
            }
            catch (Exception e)
            {

                throw;
            }

        }

        public Task<MemoryStream> TranslateExcelByStream(MemoryStream stream, string fileName)
        {
            throw new NotImplementedException();
        }
    }
}
