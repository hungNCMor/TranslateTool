
using Microsoft.Extensions.Logging;
using NPOI.HPSF;
using Spire.Presentation;
using TranslateLib.Interface;

namespace TranslateLib.PPT
{
    public class TranslatePPTWithSpire : ITranslateFile
    {
        ILogger _logger;
        ITranslate _translate;
        public TranslatePPTWithSpire(ILogger logger, ITranslate translate)
        {
            _logger = logger;
            _translate = translate;

        }
        //public async Task ReadPPTSpire()
        //    {
        //        // Assuming you have a Presentation object called 'presentation'
        //        Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

        //        // Load the PowerPoint file
        //        presentation.LoadFromFile(@"C:\Users\admin\Downloads\Tình trạng phát triển health management app.pptx");
        //        foreach (ISlide slide in presentation.Slides)
        //        {
        //            // Read the shapes within the slide
        //            foreach (IShape shape in slide.Shapes)
        //            {
        //                string shapeText = string.Empty;

        //                if (shape is IAutoShape)
        //                {
        //                    IAutoShape autoShape = (IAutoShape)shape;
        //                    shapeText = autoShape.TextFrame.Text.Trim();
        //                    if (!string.IsNullOrEmpty(shapeText))
        //                    {
        //                        autoShape.TextFrame.Text = TranslateText(shapeText.Replace("。", ". "), "en", "ja");
        //                        autoShape.TextFrame.TextRange.FontHeight = 12;

        //                        if (autoShape.TextFrame.Paragraphs.Count > 0)
        //                        {
        //                            foreach (TextParagraph paragraph in autoShape.TextFrame.Paragraphs)
        //                            {
        //                                foreach (Spire.Presentation.TextRange item in paragraph.TextRanges)
        //                                {

        //                                    //paragraph.TextRanges[0].FontHeight = 12; // Replace 18 with the desired font size
        //                                    item.FontHeight = 12; // Replace 18 with the desired font size

        //                                }
        //                            }

        //                        }
        //                    }
        //                }
        //                else if (shape is Spire.Presentation.ITable)
        //                {
        //                    var table = (Spire.Presentation.ITable)shape;
        //                    // Read the table data
        //                    if (table != null)
        //                    {
        //                        int rowCount = table.TableRows.Count;
        //                        int columnCount = table.ColumnsList.Count;

        //                        for (int row = 0; row < rowCount; row++)
        //                        {
        //                            for (int column = 0; column < columnCount; column++)
        //                            {
        //                                Spire.Presentation.Cell cell = table[column, row];
        //                                string cellText = cell.TextFrame.Text.Trim();
        //                                if (!string.IsNullOrEmpty(cellText))
        //                                    cell.TextFrame.Text = TranslateText(cellText.Replace("。", ". "), "en", "ja");
        //                                // Set the font size
        //                                cell.TextFrame.TextRange.FontHeight = 12;
        //                                // Perform operations with cellText
        //                            }
        //                        }
        //                    }
        //                }
        //            }

        //        }
        //        presentation.SaveToFile(@"C:\Users\admin\Downloads\Tình trạng phát triển health management app2.pptx", FileFormat.Pptx2013);
        //        //// Assuming you have a Slide object called 'slide'
        //        //ISlide slide = presentation.Slides[0]; // Example slide index

        //        //// Assuming you have a Shape object called 'shape'
        //        //IShape shape = slide.Shapes[0]; // Example shape index

        //        //// Read the shape properties
        //        //string shapeText = shape.AlternativeText.Trim();

        //        // Close the presentation
        //        presentation.Dispose();
        //        return Ok();
        //    }

        public async Task<MemoryStream> TranslateFileByPath(string path)
        {
            // Assuming you have a Presentation object called 'presentation'
            Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

            // Load the PowerPoint file
            presentation.LoadFromFile(path);
            var tasks = new List<Task>();
            foreach (ISlide slide in presentation.Slides)
            {
                tasks.Add(HandleSlide(slide));
            }
            await Task.WhenAll(tasks);
            return (MemoryStream)presentation.GetStream();
        }
        private async Task HandleSlide(ISlide slide)
        {
            try
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
                            autoShape.TextFrame.Text = _translate.TranslateText(shapeText.Replace("。", ". "), "ja", "en");
                            autoShape.TextFrame.TextRange.FontHeight = 12;

                            if (autoShape.TextFrame.Paragraphs.Count > 0)
                            {
                                foreach (TextParagraph paragraph in autoShape.TextFrame.Paragraphs)
                                {
                                    foreach (Spire.Presentation.TextRange item in paragraph.TextRanges)
                                    {
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
                                        cell.TextFrame.Text = _translate.TranslateText(cellText.Replace("。", ". "), "ja", "en");
                                    // Set the font size
                                    cell.TextFrame.TextRange.FontHeight = 12;
                                    // Perform operations with cellText
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception x)
            {

                throw;
            }

        }
        //public async Task TranslateFileByPathSavePath(string path)
        //{
        //    // Assuming you have a Presentation object called 'presentation'
        //    Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

        //    // Load the PowerPoint file
        //    presentation.LoadFromFile(path);
        //    var tasks = new List<Task>();
        //    foreach (ISlide slide in presentation.Slides)
        //    {
        //        tasks.Add(HandleSlide(slide));
        //    }
        //    await Task.WhenAll(tasks);
        //    var newPath = path.Replace(".pptx", "2.pptx");
        //    presentation.SaveToFile(newPath, FileFormat.Pptx2013);
        //}
        public async Task TranslateFileByPathSavePath(string path)
        {
            // Assuming you have a Presentation object called 'presentation'
            Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

            // Load the PowerPoint file
            presentation.LoadFromFile(path);
            var tasks = new List<Task>();
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
                            autoShape.TextFrame.Text = _translate.TranslateText(shapeText.Replace("。", ". "), "ja", "en");
                            autoShape.TextFrame.TextRange.FontHeight = 12;

                            if (autoShape.TextFrame.Paragraphs.Count > 0)
                            {
                                foreach (TextParagraph paragraph in autoShape.TextFrame.Paragraphs)
                                {
                                    foreach (Spire.Presentation.TextRange item in paragraph.TextRanges)
                                    {
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
                                        cell.TextFrame.Text = _translate.TranslateText(cellText.Replace("。", ". "), "ja", "en");
                                    // Set the font size
                                    cell.TextFrame.TextRange.FontHeight = 12;
                                    // Perform operations with cellText
                                }
                            }
                        }
                    }
                }
            }
            var newPath = path.Replace(".pptx", "2.pptx");
            presentation.SaveToFile(newPath, FileFormat.Pptx2013);
        }

        public Task<MemoryStream> TranslateFileByStream(MemoryStream stream, string fileName)
        {
            throw new NotImplementedException();
        }
    }
}


