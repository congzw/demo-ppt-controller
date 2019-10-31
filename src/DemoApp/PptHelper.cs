using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Core;
using OfficeCtrlDemo;
using PPT = Microsoft.Office.Interop.PowerPoint;

namespace DemoApp
{
    public class SlideInfo
    {
        public string Name { get; set; }
        public int SlideId { get; set; }
        public int PrintSteps { get; set; }

        public int SlideNumber { get; set; }
        public int SlideIndex { get; set; }

        public int SectionNumber { get; set; }
        public int SectionIndex { get; set; }
    }

    public class PptInfo
    {
        public PptInfo()
        {
            Slides = new List<SlideInfo>();
        }

        public string Path { get; set; }

        public IList<SlideInfo> Slides { get; set; }

        public override string ToString()
        {
            var stringBuilder = new StringBuilder();
            foreach (var slideInfo in Slides)
            {
                stringBuilder.AppendLine(slideInfo.AsIniString(new []{ "SectionNumber", "SectionIndex" }));
            }

            stringBuilder.AppendLine(this.AsIniString(new string[]{ "Slides" }));
            return stringBuilder.ToString();
        }
    }

    public class PptHelper : IDisposable
    {
        public PptHelper()
        {
            _app = new PPT.Application();
        }

        private PPT.Application _app = null;
        private PPT.Presentation _presentation = null;
        
        public void Open(string filePath)
        {
            _presentation = _app.Presentations.Open(filePath,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse,
                MsoTriState.msoFalse);
            var slideShowSettings = _presentation.SlideShowSettings;
            slideShowSettings.Run();
        }
        
        public PptInfo GetPptInfo()
        {
            var pptInfo = new PptInfo();
            if (_presentation == null)
            {
                return pptInfo;
            }

            pptInfo.Path = _presentation.Path;
            foreach (PPT.Slide slid in _presentation.Slides)
            {
                var slideInfo = new SlideInfo();

                slideInfo.Name = slid.Name;
                slideInfo.SlideId = slid.SlideID;
                slideInfo.PrintSteps = slid.PrintSteps;
                slideInfo.SlideNumber = slid.SlideNumber;
                slideInfo.SlideIndex = slid.SlideIndex;
                slideInfo.SectionNumber = slid.SectionNumber;
                slideInfo.SectionIndex = slid.sectionIndex;

                pptInfo.Slides.Add(slideInfo);
            }

            return pptInfo;
        }

        public void Next()
        {
            _presentation?.SlideShowWindow?.View?.Next();
        }

        public void Pre()
        {
            _presentation?.SlideShowWindow?.View?.Previous();
        }
        
        public void Close()
        {
            //_presentation?.SlideShowWindow?.View?.Exit();
            _presentation?.Close();
            _presentation = null;
        }

        public void Dispose()
        {
            _presentation?.Close();
            _app?.Quit();
            _app = null;
            _presentation = null;
        }
    }
}
