using Microsoft.Office.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;

namespace SampleDocumentCreator
{
    public interface IDocument : IDisposable
    {
        string Title { get; set; }
        string Extract { get; set; }
        string FileName { get; }
        string FullPath { get; }
        string SaveArticleToFile();
        void GetRandomArticle();
    }

    public class Document
    {
        public string Title { get; set; } = string.Empty;
        public string Extract { get; set; } = string.Empty;
        public string FileName { get; internal set; } = string.Empty;
        public string FullPath { get { return $"{Environment.CurrentDirectory}\\{FileName}"; } }

        internal object _missing = System.Reflection.Missing.Value;

        internal static Document DownloadArticle()
        {
            var url = "https://en.wikipedia.org/w/api.php?format=json&action=query&generator=random&grnnamespace=0&prop=extracts&explaintext=1";
            using (var client = new HttpClient())
            {
                var response = client.GetAsync(url).Result;
                if (!response.IsSuccessStatusCode) throw new Exception(response.StatusCode.ToString());
                var content = response.Content.ReadAsStringAsync().Result;

                var o = JObject.Parse(content);
                var title = GetPropertyValue(o, "title");

                var extract = GetPropertyValue(o, "extract");
                if (extract.IndexOf("==") > 0) extract = extract.Substring(0, extract.IndexOf("=="));
                extract = extract.Trim();

                return new Document()
                {
                    Title = title,
                    Extract = extract
                };
            }
        }

        private static string GetPropertyValue(JObject o, string name)
        {
            var titleIoken = o.SelectToken($"$.query.pages..{name}");
            return titleIoken.ToString();
        }
    }

    public class WordDocument : Document, IDocument
    {
        private Microsoft.Office.Interop.Word.Application _word;

        public WordDocument()
        {
            _word = new Microsoft.Office.Interop.Word.Application();
            _word.Visible = false;
        }

        public void Dispose()
        {
            foreach (Microsoft.Office.Interop.Word.Document d in _word.Documents) { d.Close(); }
            _word.Quit();
            _word = null;
        }

        public void GetRandomArticle()
        {
            var minExtractLength = 800;
            while (Extract.Length < minExtractLength)
            {
                var a = Document.DownloadArticle();
                this.Extract = a.Extract;
                this.Title = a.Title;
            }
            return;
        }

        public string SaveArticleToFile()
        {
            var doc = GenerateDocument();
            FileName = $"{Title}.docx";
            doc.SaveAs2(FullPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Microsoft.Office.Interop.Word.WdCompatibilityMode.wdWord2013);
            Console.WriteLine($"Saved {FileName}..");
            return FileName;
        }

        private Microsoft.Office.Interop.Word.Document GenerateDocument()
        {
            var doc = _word.Documents.Add(ref _missing, ref _missing, ref _missing, ref _missing);

            var p1 = doc.Content.Paragraphs.Add(ref _missing);
            p1.Range.Text = "";
            p1.Range.set_Style(Microsoft.Office.Interop.Word.WdBuiltinStyle.wdStyleTitle);
            p1.Range.Text = Title;
            p1.Range.InsertParagraphAfter();

            var p2 = doc.Content.Paragraphs.Add(ref _missing);
            p2.Range.Text = Extract;
            p2.Range.InsertParagraphAfter();
            return doc;
        }
    }

    public class ExcelDocument : Document, IDocument
    {
        private Microsoft.Office.Interop.Excel.Application _excel;

        public ExcelDocument()
        {
            _excel = new Microsoft.Office.Interop.Excel.Application();
            _excel.Visible = false;
        }

        public void Dispose()
        {
            foreach (Microsoft.Office.Interop.Excel.Workbook w in _excel.Workbooks) { w.Close(false, null, null); }
            _excel.Quit();
            _excel = null;
        }

        public void GetRandomArticle()
        {
            var minExtractLength = 100;
            while (Extract.Length < minExtractLength)
            {
                var a = Document.DownloadArticle();
                this.Extract = a.Extract;
                this.Title = a.Title;
            }
            return;
        }

        public string SaveArticleToFile()
        {
            var workbook = GenerateDocument();
            FileName = $"{Title}.xlsx";
            workbook.SaveAs(FullPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, _missing, _missing, _missing, _missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, _missing, _missing, _missing, _missing, _missing);
            Console.WriteLine($"Saved {FileName}..");
            return FileName;
        }

        private Microsoft.Office.Interop.Excel.Workbook GenerateDocument()
        {
            object missing = System.Reflection.Missing.Value;
            var workbook = _excel.Workbooks.Add(missing);
            var worksheet = workbook.Worksheets.get_Item(1);

            worksheet.Cells[1, 1] = Title;

            var words = Extract.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            for (var row = 0; row < words.Length % 10; row += 1)
            {
                for (int col = 0; col < 10; col++)
                {
                    var index = (10 * row) + col;
                    if (index < words.Length) worksheet.Cells[row + 2, col + 1] = words[index];
                }
            }
            return workbook;
        }
    }

    public class PowerPointDocument : Document, IDocument
    {
        private Microsoft.Office.Interop.PowerPoint.Application _ppt;

        public PowerPointDocument()
        {
            _ppt = new Microsoft.Office.Interop.PowerPoint.Application();
        }

        public void Dispose()
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in _ppt.Presentations) { p.Close(); }
            _ppt.Quit();
        }

        public void GetRandomArticle()
        {
            var minExtractLength = 200;
            while (Extract.Length < minExtractLength)
            {
                var a = Document.DownloadArticle();
                this.Extract = a.Extract;
                this.Title = a.Title;
            }
            return;
        }

        public string SaveArticleToFile()
        {
            var presentation = GenerateDocument();
            FileName = $"{Title}.pptx";
            presentation.SaveAs(FullPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
            return FileName;
        }

        private Microsoft.Office.Interop.PowerPoint.Presentation GenerateDocument()
        {
            //var presentation = ppt.Presentations.Add(MsoTriState.msoTrue);
            var presentation = _ppt.Presentations.Add(MsoTriState.msoFalse);

            var slide = presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText]);
            var range = slide.Shapes[1].TextFrame.TextRange;
            range.Text = Title;
            range.Font.Size = 44;

            range = slide.Shapes[2].TextFrame.TextRange;
            range.Text = Extract;

            return presentation;
        }
    }
}