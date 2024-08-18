using System;
using System.Collections.Generic;

namespace SampleDocumentCreator
{
    internal class Program
    {

        static void Main(string[] args)
        {
            GenerateArticles(9);
            Console.WriteLine("Done...");
            Console.ReadKey();
        }

        static void GenerateArticles(int count)
        {
            var rnd = new Random();
            var fileNames = new List<string>();
            for (var i = 0; i < count; i++)
            {
                var r = rnd.Next(0, 3);
                switch (r)
                {
                    case 0:
                        using (var article = new Document(ArticleType.Word))
                        {
                            article.GetRandomArticle();
                            article.SaveArticleToFile();
                            fileNames.Add(article.FileName);
                        }
                        break;
                    case 1:
                        using (var article = new Document(ArticleType.Excel))
                        {
                            article.GetRandomArticle();
                            article.SaveArticleToFile();
                            fileNames.Add(article.FileName);
                        }
                        break;
                    case 2:
                        using (var article = new Document(ArticleType.PowerPoint))
                        {
                            article.GetRandomArticle();
                            article.SaveArticleToFile();
                            fileNames.Add(article.FileName);
                        }
                        break;
                }
            }

            Console.WriteLine();
            foreach (var name in fileNames) Console.WriteLine($"Completed {name}");

        }
    }
}