using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace ppedv.pocgen.Logic
{
    public class PowerPointHelper : IDisposable
    {
        public PowerPointHelper()
        {
            app = new Application();
        }
        private readonly Application app;

        public Presentation OpenPresentation(string filename) => app.Presentations.Open(filename);
        public Presentation CreateNewPresentation(string filename) => app.Presentations.Add(MsoTriState.msoFalse);
        public void SavePresentationAs(Presentation output,string filename) => output.SaveAs(filename);

        public void ExportAllSlidesAsImage(Presentation presentation,string path)
        {
            int height = Convert.ToInt32(presentation.PageSetup.SlideHeight);
            int width = Convert.ToInt32(presentation.PageSetup.SlideWidth);
            for (int i = 0; i < presentation.Slides.Count; i++)
                presentation.Slides[i + 1].Export($"{path}\\{i}.png", "PNG", width, height);
        }
        public void MergePresentationContentIntoNewPresentation(IEnumerable<string> sourceFiles, Presentation destination,int insertAtIndex)
        {
            destination.ApplyTemplate(sourceFiles.First()); // Template aus der ersten Präsentation übernehmen
            foreach (string file in sourceFiles)
                insertAtIndex += destination.Slides.InsertFromFile(file, insertAtIndex);
        }
        public void MergePresentationContentIntoNewPresentation(IEnumerable<string> sourceFiles, Presentation destination) => MergePresentationContentIntoNewPresentation(sourceFiles, destination, 0);
        public void Dispose() => app.Quit();
    }
}
