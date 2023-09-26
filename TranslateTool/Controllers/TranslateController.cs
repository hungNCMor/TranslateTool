using IronXL;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Collections;
using System.Text.Json;

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

                    // Iterate through the worksheets
                    foreach (WorkSheet sheet in worksheets)
                    {
                        // Get the name of the sheet
                        string sheetName = await TranslateText(sheet.Name);

                        sheetName = sheetName.Replace('[', '(').Replace(']', ')').Substring(0, sheetName.Length > 30 ? 30 : sheetName.Length);
                        sheet.Name = sheetName + i;
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
                                    if (sheet.GetCellAt(row, column).IsFormula)
                                    {
                                        if (String.IsNullOrEmpty(sheet.GetCellAt(row, column).Formula))
                                        {
                                            continue;
                                        }
                                        sheet.GetCellAt(row, column).Formula = sheet.GetCellAt(row, column).Formula;
                                    }
                                    else
                                    {
                                        cellValue = sheet.GetCellAt(row, column)?.StringValue;
                                        if (String.IsNullOrEmpty(cellValue))
                                        {
                                            continue;
                                        }
                                        sheet.GetCellAt(row, column).Value = TranslateText(cellValue);
                                    }
                                }
                                // Perform operations with the cell value
                                // ...
                            }
                        }
                        i++;
                        sheet.UnprotectSheet();
            
                    }
                    var fileName = await TranslateText(file.FileName);
                    return File(workBook.ToByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);

                }
            }
            catch (Exception e)
            {

                throw;
            }

        }
        public async Task HandleSheet(WorkSheet sheet, int i)
        {
            // Get the name of the sheet
            string sheetName = await TranslateText(sheet.Name);

            sheetName = sheetName.Replace('[', '(').Replace(']', ')').Substring(0, sheetName.Length > 30 ? 30 : sheetName.Length);
            sheet.Name = sheetName + i;
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
                        if (sheet.GetCellAt(row, column).IsFormula)
                        {
                            if (String.IsNullOrEmpty(sheet.GetCellAt(row, column).Formula))
                            {
                                continue;
                            }
                            sheet.GetCellAt(row, column).Formula = sheet.GetCellAt(row, column).Formula;
                        }
                        else
                        {
                            cellValue = sheet.GetCellAt(row, column)?.StringValue;
                            if (String.IsNullOrEmpty(cellValue))
                            {
                                continue;
                            }
                            sheet.GetCellAt(row, column).Value = TranslateText(cellValue);
                        }
                    }
                    // Perform operations with the cell value
                    // ...
                }
            }
            i++;
            sheet.UnprotectSheet();
        }
        private async Task<string> TranslateText(string input)
        {
            string url = String.Format
            ("https://translate.googleapis.com/translate_a/single?client=gtx&tl={0}&sl={1}&dt=t&q={2}",
             "vi", "ja", Uri.EscapeUriString(input));
            HttpClient httpClient = new HttpClient();
            string responseBody = await httpClient.GetStringAsync(url);

            // Parse the response to get the translated text
            string translatedText = ParseTranslationResponse(responseBody);

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