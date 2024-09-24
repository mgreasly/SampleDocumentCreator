using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;

namespace SampleDocumentCreator
{
    internal class PowerPointFile : File, IFile
    {
        private Application _ppt;
        private Presentation _presentation;
        public int MinLength => 200;

        public PowerPointFile()
        {
            if (_ppt == null)
            {
                _ppt = new Application();
                _ppt.Visible = MsoTriState.msoTrue;
            }
        }

        public void Dispose()
        {
            if (_presentation != null)
            {
                _presentation.Close();
                _presentation = null;
            }
            if (_ppt != null)
            {
                foreach (Presentation p in _ppt.Presentations) { p.Close(); }
                _ppt.Quit();
                _ppt = null;
            }
        }

        public void GenerateDocument()
        {
            _presentation = _ppt.Presentations.Add(MsoTriState.msoFalse);

            var slide = _presentation.Slides.AddSlide(1, _presentation.SlideMaster.CustomLayouts[PpSlideLayout.ppLayoutText]);
            var range = slide.Shapes[1].TextFrame.TextRange;
            range.Text = ArticleExtract.Title;
            range.Font.Size = 44;

            range = slide.Shapes[2].TextFrame.TextRange;
            range.Text = ArticleExtract.Extract;
        }

        public void AddLinks()
        {
        }

        public string SaveArticleToFile(string path)
        {
            this.Folder = path;
            return SaveArticleToFile();
        }

        public string SaveArticleToFile()
        {
            FileName = $"{GetValidFileName(ArticleExtract.Title)}.pptx";
            try
            {
                _presentation.SaveAs(FullPath, PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);
                Console.WriteLine($"Saved {FullPath}..");
            }
            catch (Exception e)
            {
                Console.WriteLine($"{e.Source} - {e.Message}");
            }
            return FileName;
        }
    }
}