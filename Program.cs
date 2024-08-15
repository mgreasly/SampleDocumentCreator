using System;

namespace SampleDocumentCreator
{
    internal class Program
    {

        static void Main(string[] args)
        {
            GenerateArticles(1);
            Console.WriteLine("Done...");
            Console.ReadKey();
        }

        static void GenerateArticles(int count)
        {
            for (var i = 0; i < count; i++)
            {
                using (var article = new Article(ArticleType.Word))
                {
                    article.GetRandomArticle(800);
                    article.WriteArticle(article);
                    Console.WriteLine($"{i + 1}\tSaved {article.FileName}");
                }
            }
        }
    }
}