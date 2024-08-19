using System;

namespace SampleDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            GenerateArticles(500);
            Console.WriteLine("Done...");
            Console.ReadKey();
        }

        static void GenerateArticles(int count)
        {
            var rnd = new Random();
            for (var i = 0; i < count; i++)
            {
                var r = rnd.Next(0, 5);
                switch (r)
                {
                    case 0: GeneratePowerPointFile(); break;
                    case 1: GenerateExcelFile(); break;
                    case 2: GenerateExcelFile(); break;
                    default: GenerateWordFile(); break;
                }
            }
        }

        static void GeneratePowerPointFile()
        {
            var article = ArticleExtract.DownloadWikiArticle();
            while (article.Extract.Length < 200) { article = ArticleExtract.DownloadWikiArticle(); }
            Console.WriteLine($"PPTX: Using {article.Title}");
            using (var file = new PowerPointFile())
            {
                file.ArticleExtract = article;
                file.GenerateDocument();
                file.AddLinks();
                file.SaveArticleToFile();
            }
        }

        static void GenerateExcelFile()
        {
            var article = ArticleExtract.DownloadWikiArticle();
            while (article.Extract.Length < 400) { article = ArticleExtract.DownloadWikiArticle(); }
            Console.WriteLine($"XLSX: Using {article.Title}");
            using (var file = new ExcelFile())
            {
                file.ArticleExtract = article;
                file.GenerateDocument();
                file.AddLinks();
                file.SaveArticleToFile();
            }
        }

        static void GenerateWordFile()
        {
            var article = ArticleExtract.DownloadWikiArticle();
            while (article.Extract.Length < 800) { article = ArticleExtract.DownloadWikiArticle(); }
            Console.WriteLine($"DOCX: Using {article.Title}");
            using (var file = new WordFile())
            {
                file.ArticleExtract = article;
                file.GenerateDocument();
                file.AddLinks();
                file.SaveArticleToFile();
            }
        }
    }
}