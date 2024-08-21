using Microsoft.Office.Core;
using Newtonsoft.Json.Linq;
using System;
using System.Net.Http;
using System.Security.Policy;

namespace SampleDocumentCreator
{
    public interface IFile : IDisposable
    {
        ArticleExtract ArticleExtract { get; set; }
        string FileName { get; }
        string FullPath { get; }
        int MinLength { get; }
        void GenerateDocument();
        void AddLinks();
        string SaveArticleToFile();
    }

    public class File
    {
        public ArticleExtract ArticleExtract { get; set; } = new ArticleExtract();
        public string Title { get; set; } = string.Empty;
        public string Extract { get; set; } = string.Empty;
        public string FileName { get; internal set; } = string.Empty;
        public string FullPath { get { return $"{Environment.CurrentDirectory}\\{FileName}"; } }

        internal object _missing = System.Reflection.Missing.Value;

        internal string GetValidFileName(string name)
        {
            string f = name;
            foreach (char c in System.IO.Path.GetInvalidFileNameChars()) { f = f.Replace(c, '_'); }
            return f;
        }
    }
}