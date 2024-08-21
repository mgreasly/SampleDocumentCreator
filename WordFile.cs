using Microsoft.Office.Interop.Word;
using System;

namespace SampleDocumentCreator
{
    public class WordFile : File, IFile
    {
        private Application _word;
        private Document _doc;
        public int MinLength => 800;

        public WordFile()
        {
            if (_word == null)
            {
                _word = new Application();
                _word.Visible = false;
            }
            if (_doc == null) _doc = _word.Documents.Add(ref _missing, ref _missing, ref _missing, ref _missing);
        }

        public void Dispose()
        {
            if (_doc != null)
            {
                _doc.Close();
                _doc = null;
            }
            foreach (Document d in _word.Documents) { d.Close(); }
            _word.Quit();
            _word = null;
        }

        public void GenerateDocument()
        {
            var p1 = _doc.Content.Paragraphs.Add(ref _missing);
            p1.Range.Text = "";
            p1.Range.set_Style(WdBuiltinStyle.wdStyleTitle);
            p1.Range.Text = ArticleExtract.Title;
            p1.Range.InsertParagraphAfter();

            var p2 = _doc.Content.Paragraphs.Add(ref _missing);
            p2.Range.Text = ArticleExtract.Extract;
            p2.Range.InsertParagraphAfter();
        }

        public void AddLinks()
        {
            var rnd = new Random();
            var wordsInTitle = ArticleExtract.Title.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Length;
            var count = rnd.Next(8);
            Console.WriteLine($"Adding {count} links");
            for (int i = 0; i < count; i++)
            {
                var index = rnd.Next(wordsInTitle, _doc.Words.Count + 1);
                object start = _doc.Words[index].Start;
                object end = _doc.Words[index].End;
                var range = _doc.Range(ref start, ref end);
                object address = LinkGenerator.RandomLink();
                var link = _doc.Hyperlinks.Add(range, ref address);
            }
        }

        public string SaveArticleToFile()
        {
            FileName = $"{GetValidFileName(ArticleExtract.Title)}.docx";
            try
            {
                _doc.SaveAs2(FullPath, WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: WdCompatibilityMode.wdWord2013);
                Console.WriteLine($"Saved {FullPath}..");
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e.Source} - {e.Message}");
            }
            return FileName;
        }
    }
}