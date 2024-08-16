using Newtonsoft.Json.Linq;
using System;
using System.Linq;
using System.Net.Http;

namespace SampleDocumentCreator
{
    public class Document : IDisposable
    {
        public ArticleType ArticleType { get; set; } = ArticleType.Unknown;
        public string Title { get; set; } = string.Empty;
        public string Extract { get; set; } = string.Empty;
        public string FileName { get; private set; } = string.Empty;
        public string FullPath { get { return $"{Environment.CurrentDirectory}\\{FileName}"; } }

        object _missing = System.Reflection.Missing.Value;
        private Microsoft.Office.Interop.Word.Application _word;
        private Microsoft.Office.Interop.Excel.Application _excel;

        public Document(ArticleType type)
        {
            ArticleType = type;
            switch (ArticleType)
            {
                case ArticleType.Excel:
                    if (_excel == null)
                    {
                        _excel = new Microsoft.Office.Interop.Excel.Application();
                        _excel.Visible = false;
                    }
                    break;
                case ArticleType.Word:
                    if (_word == null)
                    {
                        _word = new Microsoft.Office.Interop.Word.Application();
                        _word.Visible = false;
                    }
                    break;
                default:
                    // no action
                    break;
            }
        }
        public void Dispose()
        {
            if (_excel != null)
            {
                foreach (Microsoft.Office.Interop.Excel.Workbook w in _excel.Workbooks) { w.Close(false, null, null); }
                _excel.Quit();
                _excel = null;
            }
            if (_word != null)
            {
                foreach (Microsoft.Office.Interop.Word.Document d in _word.Documents) { d.Close(); }
                _word.Quit();
                _word = null;
            }
        }

        public string SaveArticleToFile()
        {
            switch (ArticleType)
            {
                case ArticleType.Excel:
                    var workbook = GenerateExcelDocument(_excel);
                    FileName = $"{Title}.xlsx";
                    workbook.SaveAs(FullPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook, _missing, _missing, _missing, _missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, _missing, _missing, _missing, _missing, _missing);
                    Console.WriteLine($"Saved {FileName}..");
                    return FileName;
                case ArticleType.Word:
                    var doc = GenerateWordDocument(_word);
                    FileName = $"{Title}.docx";
                    doc.SaveAs2(FullPath, Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: Microsoft.Office.Interop.Word.WdCompatibilityMode.wdWord2013);
                    Console.WriteLine($"Saved {FileName}..");
                    return FileName;
                default:
                    return string.Empty;
            }
        }
        private Microsoft.Office.Interop.Word.Document GenerateWordDocument(Microsoft.Office.Interop.Word.Application word)
        {
            var doc = word.Documents.Add(ref _missing, ref _missing, ref _missing, ref _missing);

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
        private Microsoft.Office.Interop.Excel.Workbook GenerateExcelDocument(Microsoft.Office.Interop.Excel.Application excel)
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
                    worksheet.Cells[row + 2, col + 1] = words[(10 * row) + col];
                }
            }
            return workbook;
        }

        public void GetRandomArticle()
        {
            var minExtractLength = 0;
            switch (ArticleType)
            {
                case ArticleType.Excel: minExtractLength = 100; break;
                case ArticleType.Word: minExtractLength = 800; break;
            }
            var url = "https://en.wikipedia.org/w/api.php?format=json&action=query&generator=random&grnnamespace=0&prop=extracts&explaintext=1";
            while (Extract.Length < minExtractLength)
            {
                if (Extract.Length > 0) Console.WriteLine("extract too short...");
                DownloadArticle(url);
            }
            Console.WriteLine("article found...");
            return;
        }
        private void DownloadArticle(string url)
        {
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

                Title = title;
                Extract = extract;
                Console.Write($"Downloaded {Title} ");
                return;
            }
        }
        private static string GetPropertyValue(JObject o, string name)
        {
            var titleIoken = o.SelectToken($"$.query.pages..{name}");
            return titleIoken.ToString();
        }
    }

    public enum ArticleType
    {
        Unknown = 0,
        Excel,
        Word
    }
}