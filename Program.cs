using System;
using System.IO;

namespace SampleDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var path = args[0];
            path = Path.Combine(path, string.Format("{0}-{1:00}-{2:00}", DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day));
            Directory.CreateDirectory(path);

            var count = Convert.ToInt32(args[1]);

            GenerateArticles(count, path);

            Console.WriteLine("Done...");
            Console.ReadKey();
        }

        static void GenerateArticles(int count, string path)
        {
            var rnd = new Random();
            for (var i = 0; i < count; i++)
            {
                var r = rnd.Next(0, 5);
                switch (r)
                {
                    //case 0: using (var file = new WordFile()) { ProcessArticle(file); }; break;
                    //case 1: using (var file = new ExcelFile()) { ProcessArticle(file); }; break;
                    //case 2: using (var file = new PowerPointFile()) { ProcessArticle(file); }; break;
                    default: using (var file = new WordFile()) { ProcessArticle(file); }; break;
                }
            }
        }

        static void ProcessArticle(IFile file)
        {
            var article = ArticleExtract.DownloadWikiArticle();
            while (article.Extract.Length < file.MinLength) { article = ArticleExtract.DownloadWikiArticle(); }
            Console.WriteLine($"{file.MinLength}: Using {article.Title}");

            file.ArticleExtract = article;
            file.GenerateDocument();
            file.AddLinks();
            file.SaveArticleToFile(file.FullPath);
        }
    }
}