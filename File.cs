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
    /*
        public class PowerPointDocument : Document, IDocument
        {
            private Microsoft.Office.Interop.PowerPoint.Application _ppt;

            public PowerPointDocument()
            {
            }

            public void Dispose()
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Presentation p in _ppt.Presentations) { p.Close(); }
                _ppt.Quit();
            }

            public void GetRandomArticle()
            {
                var minExtractLength = 200;
                while (Extract.Length < minExtractLength)
                {
                    var a = Document.DownloadArticle();
                    this.Extract = a.Extract;
                    this.Title = a.Title;
                }
                return;
            }

            public string SaveArticleToFile()
            {
                var presentation = GenerateDocument();
                FileName = $"{Title}.pptx";
                presentation.SaveAs(FullPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                return FileName;
            }

            private Microsoft.Office.Interop.PowerPoint.Presentation GenerateDocument()
            {
                //var presentation = ppt.Presentations.Add(MsoTriState.msoTrue);
                var presentation = _ppt.Presentations.Add(MsoTriState.msoFalse);

                var slide = presentation.Slides.AddSlide(1, presentation.SlideMaster.CustomLayouts[Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText]);
                var range = slide.Shapes[1].TextFrame.TextRange;
                range.Text = Title;
                range.Font.Size = 44;

                range = slide.Shapes[2].TextFrame.TextRange;
                range.Text = Extract;

                return presentation;
            }
        }
    */
}