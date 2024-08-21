﻿using System;

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
                file.SaveArticleToFile();
            }
        }
    }
}