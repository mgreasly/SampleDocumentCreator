using System;
using System.IO;

namespace SampleDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = args[0];
            GenerateArticles(2, path);
            Console.WriteLine("Done...");
            Console.ReadKey();
        }

        static void GenerateArticles(int count, string path)
        {
            var rnd = new Random();
            for (var i = 0; i < count; i++)
            {
                var r = rnd.Next(0, 5);
                IFile file = null;
                switch (r)
                {
                    //case 0: file = new PowerPointFile(); break;
                    //case 1: file = new ExcelFile(); break;
                    //case 2: file = new ExcelFile(); break;
                    default: file = new WordFile(); break;
                }
                var article = ArticleExtract.DownloadWikiArticle();
                while (article.Extract.Length < file.MinLength) { article = ArticleExtract.DownloadWikiArticle(); }
                Console.WriteLine($"{file.MinLength}: Using {article.Title}");

                file.ArticleExtract = article;
                file.GenerateDocument();
                file.AddLinks();
                file.SaveArticleToFile(path);
            }
        }
    }
}