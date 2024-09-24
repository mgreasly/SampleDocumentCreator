using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;

namespace SampleDocumentCreator
{
    internal class ExcelFile : File, IFile
    {
        private Application _excel;
        private Workbook _wkbk;
        private Worksheet _wkst;
        public int MinLength => 400;

        public ExcelFile()
        {
            if (_excel == null)
            {
                _excel = new Application();
                _excel.Visible = false;
            }
        }

        public void Dispose()
        {
            foreach (Workbook w in _excel.Workbooks) { w.Close(false, null, null); }
            _excel.Quit();
            _excel = null;
        }

        public void GenerateDocument()
        {
            _wkbk = _excel.Workbooks.Add(_missing);
            _wkst = _wkbk.Worksheets.get_Item(1);

            _wkst.Cells[1, 1] = ArticleExtract.Title;

            var words = ArticleExtract.Extract.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            for (var i = 0; i < words.Length; i++)
            {
                var row = GetRow(i);
                var col = GetCol(i);
                _wkst.Cells[row + 2, col + 1] = words[i];
            }
        }

        public void AddLinks()
        {
            var rnd = new Random();
            var count = rnd.Next(4);
            var words = ArticleExtract.Extract.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);

            Console.WriteLine($"Adding {count} links");
            for (int i = 0; i < count; i++)
            {
                var index = rnd.Next(words.Count());
                var row = GetRow(index);
                var col = GetCol(index);
                var range = _wkst.Cells[row + 2, col + 1];
                string address = LinkGenerator.RandomLink();
                var link = _wkst.Hyperlinks.Add(range, address, TextToDisplay: words[index]);
            }
        }

        public string SaveArticleToFile(string path)
        {
            this.Folder = path;
            return SaveArticleToFile();
        }

        public string SaveArticleToFile()
        {
            FileName = $"{GetValidFileName(ArticleExtract.Title)}.xlsx";
            try
            {
                _wkbk.SaveAs(FullPath, XlFileFormat.xlOpenXMLWorkbook, _missing, _missing, _missing, _missing, XlSaveAsAccessMode.xlExclusive, _missing, _missing, _missing, _missing, _missing);
                Console.WriteLine($"Saved {FullPath}..");
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e.Source} - {e.Message}");
            }
            return FileName;
        }

        private static int GetRow(int count) => (count / 10) + 1;
        private static int GetCol(int count) => (count % 10) + 1;
    }
}