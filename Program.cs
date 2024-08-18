using System;
using System.Collections.Generic;

namespace SampleDocumentCreator
{
    internal class Program
    {
        static void Main(string[] args)
        {
            GenerateArticles(5);
            Console.WriteLine("Done...");
            Console.ReadKey();
        }

        static void GenerateArticles(int count)
        {
            var rnd = new Random();
            for (var i = 0; i < count; i++)
            {
                var r = rnd.Next(0, 3);
                switch(r){
                    case 0:
                        using (var article = new ExcelDocument())
                        {
                            article.GetRandomArticle();
                            article.SaveArticleToFile();
                        }
                        break;
                    case 1:
                        using (var article = new PowerPointDocument())
                        {
                            article.GetRandomArticle();
                            article.SaveArticleToFile();
                        }
                        break;
                    default:
                        using (var article = new WordDocument())
                        {
                            article.GetRandomArticle();
                            article.SaveArticleToFile();
                        }
                        break;
                }
            }
        }
    }
}