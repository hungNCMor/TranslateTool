using BitMiracle.LibTiff.Classic;
using IronXL;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.ViewFeatures;
using Microsoft.Office.Core;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System.Collections;
using System.ComponentModel;
using System.Data.Common;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text.Json;
using System.Threading.Tasks;
using Spire.Presentation;
using Microsoft.Office.Interop.PowerPoint;
using Spire.Pdf.Exporting.XPS.Schema;
using IronXL.Drawing;

namespace TranslateTool.Controllers
{
    [ApiController]
    [Route("[controller]/[action]")]
    public class TranslateController : ControllerBase
    {

        private readonly ILogger<TranslateController> _logger;

        public TranslateController(ILogger<TranslateController> logger)
        {
            _logger = logger;
        }

        [HttpPost("Translate")]
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
        private async Task TranslateFile(string path)
        {
            try
            {
                var t = DateTime.Now;
                var fileName = path.Replace(".xlsx", "").Split('\\').Last();
                Console.WriteLine($" {fileName} ThreadId: {System.Environment.CurrentManagedThreadId}");
                var path2 = path.Replace(".xlsx", "").Replace(fileName, fileName + "2.xlsx");
                using (var ms = new MemoryStream())
                {
                    WorkBook workBook = WorkBook.Load(path);

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
                    var t2 = DateTime.Now;
                    //UpdateWorkBook(workBook, tasks);
                    _logger.LogInformation("TranslateText time " + (DateTime.Now - t2));
                    _logger.LogInformation("total time " + (DateTime.Now - t));
                    workBook.SaveAs(path2);
                }
            }
            catch (Exception)
            {

                throw;
            }


        }
        [HttpPost("TranslateFolder")]
        public async Task TranslateFolder(string path)
        {
            DirectoryInfo d = new DirectoryInfo(path);
            var listTasks = new List<Task>();
            var files = d.GetFiles();
            try
            {
                //foreach (var file in files)
                //{
                //    listTasks.Add(TranslateFile(file.FullName));
                //}
                //await Task.WhenAll(listTasks);
                int maxWorkerThreads = 20; // Set the maximum number of worker threads
                int maxCompletionPortThreads = 20; // Set the maximum number of IO completion port threads
                ThreadPool.SetMaxThreads(maxWorkerThreads, maxCompletionPortThreads);
                await Task.Run(() => Parallel.ForEach(files, new ParallelOptions { MaxDegreeOfParallelism = 5 },
                                    x => TranslateFile(x.FullName)
                                    ));

            }
            catch (Exception cc)
            {
                throw;
            }

        }
        [HttpPost("RenameFolder")]
        public async Task<ActionResult> RenameFolder(string path)
        {
            //var nameFolder = path.Split('\\').LastOrDefault();
            //var path1 = String.Join('\\', path.Split('\\').Take(path.Split('\\').Length - 1));
            //var fullPath = Path.Combine(path1, nameFolder + TranslateText(nameFolder));
            //Directory.Move(path, fullPath);
            //RenameFolderAndFile(fullPath);
            RenameFolderAndFile(path);
            return Ok();
        }
        [HttpPost("ReNameFolderToOld")]
        public async Task<ActionResult> ReNameFolderToOld(string path)
        {
            //var nameFolder = path.Split('\\').LastOrDefault();
            //var path1 = String.Join('\\', path.Split('\\').Take(path.Split('\\').Length - 1));
            //var fullPath = Path.Combine(path1, nameFolder + TranslateText(nameFolder));
            //Directory.Move(path, fullPath);
            //RenameFolderAndFile(fullPath);
            RenameFolderAndFileTOld(path);
            return Ok();
        }
        private void RenameFolderAndFile(string path)
        {
            DirectoryInfo d = new DirectoryInfo(path);
            var folders = d.GetDirectories();
            var files = d.GetFiles();
            foreach (var folder in folders)
            {
                var newPath = System.IO.Path.Combine(path, folder.Name + TranslateText(folder.Name));
                if (newPath != folder.FullName)
                    Directory.Move(folder.FullName, newPath);
            }
            foreach (var file in files)
            {
                var newNameFile = System.IO.Path.Combine(path, file.Name + TranslateText(file.Name));
                newNameFile = newNameFile.Replace("/", ".");
                if (newNameFile != file.FullName)
                    System.IO.File.Move(file.FullName, newNameFile);
            }
        }
        private void RenameFolderAndFileTOld(string path)
        {
            DirectoryInfo d = new DirectoryInfo(path);
            var files = d.GetFiles();

            foreach (var file in files)
            {
                var newNameFile = file.Name.Split(".xlsx").FirstOrDefault() + ".xlsx";
                var newFullName = System.IO.Path.Combine(path, newNameFile);
                if (newFullName != file.FullName)
                    System.IO.File.Move(file.FullName, newNameFile);
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
            try
            {
                string sheetName = TranslateText(sheet.Name).Replace("/", ".").Replace("・", ".");

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
                                //    if (String.IsNullOrEmpty(sheet.GetCellAt(row, column).Formula))
                                //    {
                                //        continue;
                                //    }
                                //    //result.Add(new DataResult { column = column, row = row, Value = sheet.GetCellAt(row, column).Formula, IsFormula = true, Sheet = i });
                                //    //                            sheet.GetCellAt(row, column).Formula = sheet.GetCellAt(row, column).Formula;
                                //}
                                //else
                                //{
                                _logger.LogInformation($"GetCellAt row={row}, column={column}, sheet ={sheet.Name}");
                                cellValue = sheet.GetCellAt(row, column)?.StringValue;

                                if (String.IsNullOrEmpty(cellValue?.Trim()))
                                {
                                    continue;
                                }
                                _logger.LogInformation($"cellValue={cellValue}");
                                //result.Add(new DataResult { column = column, row = row, Value = cellValue, IsFormula = false, Sheet = i });
                                sheet.GetCellAt(row, column).Value = TranslateText(cellValue).Replace("/", ".");
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
            catch (Exception x)
            {

                throw x;
            }

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
        private string TranslateText(string input, string des = "vi", string src = "ja")
        {
            var t2 = DateTime.Now;
            string url = String.Format
            ("https://translate.googleapis.com/translate_a/single?client=gtx&tl={0}&sl={1}&dt=t&q={2}",
             des, src, Uri.EscapeUriString(input));
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
            try
            {
                string extractedString = "";
                if (data[0].Count == 1) { extractedString = data[0][0][0].ToString(); }
                else if (data[0].Count > 1)
                {
                    foreach (dynamic item in data[0])
                    {
                        //extractedString = item.ToString();
                        string extractedString2 = item[0].ToString();
                        extractedString = extractedString + extractedString2;
                    }
                }


                return extractedString;
            }
            catch (Exception c)
            {

                throw;
            }
        }
        [HttpPost("CreateExcel")]
        public void Create()
        {
            try
            {
                CreateExcel();
            }
            catch (Exception)
            {

                throw;
            }


        }
        [HttpPost("Replace")]
        public void Replace()
        {

            try
            {
                ReplaceSheetName();
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void CreateExcel()
        {
            WorkBook xlsWorkbook = WorkBook.Create(ExcelFileFormat.XLS);
            WorkSheet xlsSheet = xlsWorkbook.CreateWorkSheet("画面・機能概要");
            xlsWorkbook.SaveAs("excelfile.xlsx");
        }
        private void ReplaceSheetName()
        {
            WorkBook workBook = WorkBook.Load("excelfile.xlsx");
            WorkSheet workSheet = workBook.GetWorkSheet("画面・機能概要");
            string sheetName = workSheet.Name;
            Console.WriteLine(sheetName);
            //workSheet.Name = "画面機能概要";
            workSheet.Name = "画面機能概要";
            workBook.SaveAs("newexcelfile.xlsx");
        }
        [HttpPost("Translate")]
        public async Task<ActionResult> ReplaceSheetName2(IFormFile file)
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
                    WorkSheet workSheet = workBook.GetWorkSheet("画面・機能概要");
                    string sheetName = workSheet.Name;
                    Console.WriteLine(sheetName);
                    //workSheet.Name = "画面機能概要";
                    workSheet.Name = "画面機能概要";
                    workBook.SaveAs("newexcelfile.xlsx");
                    return File(workBook.ToByteArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "vcc");

                }
            }
            catch (Exception e)
            {

                throw;
            }

        }
        [HttpPost("ReadPPT")]
        public async Task<ActionResult> ReadPPT()
        {
            Microsoft.Office.Interop.PowerPoint.Application PowerPoint_App = new Microsoft.Office.Interop.PowerPoint.Application();
            Microsoft.Office.Interop.PowerPoint.Presentations multi_presentations = PowerPoint_App!.Presentations;
            Microsoft.Office.Interop.PowerPoint.Presentation presentation = multi_presentations.Open(@"C:\Users\admin\Downloads\Tình trạng phát triển health management app.pptx");
            string presentation_text = "";
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                foreach (var item in presentation.Slides[i + 1].Shapes)
                {
                    var shape = (Microsoft.Office.Interop.PowerPoint.Shape)item;
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            var text = textRange.Text;
                            presentation_text += text + " ";
                        }
                    }
                }
            }
            PowerPoint_App.Quit();
            return Ok();
        }



        [HttpPost("ReadPPTSpire")]
        public async Task<ActionResult> ReadPPTSpire()
        {
            // Assuming you have a Presentation object called 'presentation'
            Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

            // Load the PowerPoint file
            presentation.LoadFromFile(@"C:\Users\admin\Downloads\Tình trạng phát triển health management app.pptx");
            foreach (ISlide slide in presentation.Slides)
            {
                // Read the shapes within the slide
                foreach (IShape shape in slide.Shapes)
                {
                    string shapeText = string.Empty;

                    if (shape is IAutoShape)
                    {
                        IAutoShape autoShape = (IAutoShape)shape;
                        shapeText = autoShape.TextFrame.Text.Trim();
                        if (!string.IsNullOrEmpty(shapeText))
                        {
                            autoShape.TextFrame.Text = TranslateText(shapeText.Replace("。",". "), "en", "ja");
                            autoShape.TextFrame.TextRange.FontHeight = 12;

                            if (autoShape.TextFrame.Paragraphs.Count > 0)
                            {
                                foreach (TextParagraph paragraph in autoShape.TextFrame.Paragraphs)
                                {
                                    foreach (Spire.Presentation.TextRange item in paragraph.TextRanges)
                                    {

                                        //paragraph.TextRanges[0].FontHeight = 12; // Replace 18 with the desired font size
                                        item.FontHeight = 12; // Replace 18 with the desired font size
                                        
                                    }
                                }

                            }
                        }
                    }
                    else if (shape is Spire.Presentation.ITable)
                    {
                        var table = (Spire.Presentation.ITable)shape;
                        // Read the table data
                        if (table != null)
                        {
                            int rowCount = table.TableRows.Count;
                            int columnCount = table.ColumnsList.Count;

                            for (int row = 0; row < rowCount; row++)
                            {
                                for (int column = 0; column < columnCount; column++)
                                {
                                    Spire.Presentation.Cell cell = table[column, row];
                                    string cellText = cell.TextFrame.Text.Trim();
                                    if (!string.IsNullOrEmpty(cellText))
                                        cell.TextFrame.Text = TranslateText(cellText.Replace("。", ". "), "en", "ja");
                                    // Set the font size
                                    cell.TextFrame.TextRange.FontHeight = 12;
                                    // Perform operations with cellText
                                }
                            }
                        }
                    }
                }

            }
            presentation.SaveToFile(@"C:\Users\admin\Downloads\Tình trạng phát triển health management app2.pptx", FileFormat.Pptx2013);
            //// Assuming you have a Slide object called 'slide'
            //ISlide slide = presentation.Slides[0]; // Example slide index

            //// Assuming you have a Shape object called 'shape'
            //IShape shape = slide.Shapes[0]; // Example shape index

            //// Read the shape properties
            //string shapeText = shape.AlternativeText.Trim();

            // Close the presentation
            presentation.Dispose();
            return Ok();
        }

    }
}