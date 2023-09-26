using BitMiracle.LibTiff.Classic;
using IronXL;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Newtonsoft.Json;
using System.Collections;
using System.ComponentModel;
using System.Data.Common;
using System.Text.Json;
using System.Threading.Tasks;

namespace TranslateTool.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class TranslateController : ControllerBase
    {

        private readonly ILogger<TranslateController> _logger;

        public TranslateController(ILogger<TranslateController> logger)
        {
            _logger = logger;
        }

        [HttpPost(Name = "Translate")]
        public async Task<ActionResult> TranslateFile(IFormFile file)
        {
            try
            {
                ThreadPool.GetMaxThreads(out int vc, out int maxThreads);
                Console.WriteLine($"Current maximum thread count:{vc} {maxThreads}");
                var t = DateTime.Now;

                if (file == null)
                    throw new ArgumentNullException(nameof(file));
                using (var ms = new MemoryStream())
                {
                    await file.CopyToAsync(ms);
                    ms.Seek(0, SeekOrigin.Begin);

                    // Supported spreadsheet formats for reading include: XLSX, XLS, CSV and TSV

                    WorkBook workBook = WorkBook.Load(ms);
                    // Loop through each sheet in the workbook
                    // Get the worksheets
                    List<WorkSheet> worksheets = workBook.WorkSheets.ToList();
                    int i = 0;
                    var tasks = new List<Task<List<DataResult>>>();
                    // Iterate through the worksheets
                    foreach (WorkSheet sheet in worksheets)
                    {
                        tasks.Add(HandleWorkSheet(sheet, i));
                        i++;
                    }
                    var t1 = DateTime.Now;
                    await Task.WhenAll(tasks);
                    _logger.LogInformation("GetValue time " + (DateTime.Now - t1));
                    var fileName = TranslateText(file.FileName);
                    var t2 = DateTime.Now;
                    //UpdateWorkBook(workBook, tasks);
                    _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
                    _logger.LogInformation("total time " + (DateTime.Now - t));
                    return File(workBook.ToByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

                }
            }
            catch (Exception e)
            {

                throw;
            }

        }
        private async Task<WorkBook> UpdateWorkBook(WorkBook workBook, List<Task<List<DataResult>>> results, int maxThreads = 1000)
        {
            var num = results.Count / maxThreads == 0 ? 1 : results.Count / maxThreads;
            var listTaskTranslate = new List<Task<List<DataResult>>>();
            var lists = results.SelectMany(x => x.Result).ToList();
            for (int i = 0; i < lists.Count / num; i++)
            {
                listTaskTranslate.Add(HandleTranslate(lists.Skip(i * num).Take(num).ToList()));
            }
            await Task.WhenAll(listTaskTranslate);

            foreach (var result1 in listTaskTranslate)
            {
                var result = (result1.Result);
                foreach (var result2 in result)
                {
                    workBook.WorkSheets[result2.Sheet].GetCellAt(result2.row, result2.column).Value = result2.Value;
                }
            }
            return workBook;
        }
        private async Task<List<DataResult>> HandleWorkSheet(WorkSheet sheet, int i)
        {
            string sheetName = TranslateText(sheet.Name);

            sheetName = sheetName.Replace('[', '(').Replace(']', ')').Substring(0, sheetName.Length > 30 ? 30 : sheetName.Length);
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
                        //    if (String.IsNullOrEmpty(sheet.GetCellAt(row, column).Formula))
                        //    {
                        //        continue;
                        //    }
                        //    //result.Add(new DataResult { column = column, row = row, Value = sheet.GetCellAt(row, column).Formula, IsFormula = true, Sheet = i });
                        //    //                            sheet.GetCellAt(row, column).Formula = sheet.GetCellAt(row, column).Formula;
                        //}
                        //else
                        //{
                            cellValue = sheet.GetCellAt(row, column)?.StringValue;
                            if (String.IsNullOrEmpty(cellValue))
                            {
                                continue;
                            }
                            //result.Add(new DataResult { column = column, row = row, Value = cellValue, IsFormula = false, Sheet = i });
                            sheet.GetCellAt(row, column).Value = TranslateText(cellValue);
                        }
                    }
                    // Perform operations with the cell value
                    // ...
                }
            }
            i++;
            sheet.UnprotectSheet();
            return result;
        }
        private async Task<List<DataResult>> HandleTranslate(List<DataResult> datas)
        {
            var result = new List<DataResult>();
            foreach (DataResult dataResult in datas)
            {
                dataResult.Value = TranslateText(dataResult.Value);
                result.Add(dataResult);
            }
            return result;
        }
        public class DataResult
        {
            public string Value { get; set; }
            public bool IsFormula { get; set; } = false;
            public int row { get; set; }
            public int column { get; set; }
            public int Sheet { get; set; }
        }
        private string TranslateText(string input)
        {
            var t2 = DateTime.Now;
            string url = String.Format
            ("https://translate.googleapis.com/translate_a/single?client=gtx&tl={0}&sl={1}&dt=t&q={2}",
             "vi", "ja", Uri.EscapeUriString(input));
            HttpClient httpClient = new HttpClient();
            string responseBody = httpClient.GetStringAsync(url).Result;

            // Parse the response to get the translated text
            string translatedText = ParseTranslationResponse(responseBody);
            _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
            return translatedText;
        }
        private string ParseTranslationResponse(string response)
        {
            dynamic data = JsonConvert.DeserializeObject(response);
            string extractedString = data[0][0][0].ToString();
            return extractedString;
        }
    }
}