using Microsoft.Office.Interop.Word;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;

namespace SampleDocumentCreator
{
    public class Article:IDisposable
    {
        public ArticleType ArticleType { get; set; } = ArticleType.Unknown;
        public string Title { get; set; } = string.Empty;
        public string Extract { get; set; } = string.Empty;
        public string FileName { get; private set; } = string.Empty;
        public string FullPath
        {
            get { return $"{Environment.CurrentDirectory}//{FileName}"; }
        }

        private Application _word;

        public Article(ArticleType type)
        {
            ArticleType = type;
            if (_word == null)
            {
                _word = new Microsoft.Office.Interop.Word.Application();
                _word.Visible = false;
            }
        }
        public void Dispose()
        {
            foreach (Document d in _word.Documents)
            {
                d.Close();
            }
            _word.Quit();
            _word = null;
        }

        public string WriteArticle(Article article)
        {
            var doc = GenerateDocument(article, _word);
            article.FileName = $"{article.Title}.docx";
            doc.SaveAs2(article.FullPath, WdSaveFormat.wdFormatXMLDocument, CompatibilityMode: WdCompatibilityMode.wdWord2013);
            Console.WriteLine($"Saved {article.FileName}..");
            return article.FileName;
        }
        private Document GenerateDocument(Article article, Application word)
        {
            object missing = System.Reflection.Missing.Value;
            var doc = word.Documents.Add(ref missing, ref missing, ref missing, ref missing);

            var p1 = doc.Content.Paragraphs.Add(ref missing);
            p1.Range.Text = "";
            p1.Range.set_Style(WdBuiltinStyle.wdStyleTitle);
            p1.Range.Text = article.Title;
            p1.Range.InsertParagraphAfter();

            var p2 = doc.Content.Paragraphs.Add(ref missing);
            p2.Range.Text = article.Extract;
            p2.Range.InsertParagraphAfter();
            return doc;
        }

        public void GetRandomArticle(int minExtractLength)
        {
            var url = "https://en.wikipedia.org/w/api.php?format=json&action=query&generator=random&grnnamespace=0&prop=extracts&explaintext=1";
            while (Extract.Length < minExtractLength)
            {
                if (Extract.Length > 0) Console.WriteLine("extract too short...");
                 DownloadArticle(url);
            }
            Console.WriteLine();
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
        Unknown,
        Excel,
        Word
    }
}