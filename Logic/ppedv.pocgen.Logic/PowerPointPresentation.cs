using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Reflection;
using ppedv.pocgen.Domain.Interfaces;
using ppedv.pocgen.Domain.Models;
using System.Diagnostics;

namespace ppedv.pocgen.Logic
{
    public class PowerPointPresentation : IPowerPointPresentation
    {
        private Presentation presentation;
        public PowerPointPresentation(Presentation presentation)
        {
            this.presentation = presentation;
        }
        public Slides Slides => presentation.Slides;
        public int NumberOfSlidesInPresentation => Slides.Count;

        public void Dispose()
        {
            try
            {
                presentation?.Close();
                presentation = null;
                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PowerPointPresentation disposed");
            }
            catch (Exception)
            {
                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Exception when trying to dispose: Close() not possible");
            }
        }

        public SlideType GetSlideType(int pageNumber)
        {
            if (pageNumber > presentation.Slides.Count) // Slides[] fängt bei 1 und nicht bei 0 an !
            {
                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] page {pageNumber} nonexistent");
                throw new ArgumentException($"Angeforderte Seite {pageNumber} ist nicht vorhanden. Die Präsentation hat nur {presentation.Slides.Count} Seiten !");
            }

            switch (presentation.Slides[pageNumber].Layout)
            {
                case PpSlideLayout.ppLayoutTitle: // Erste Seite -> Titelfolie, jede andere: Modultitelfolie
                    return SlideType.Title;
                case PpSlideLayout.ppLayoutText:
                case PpSlideLayout.ppLayoutSectionHeader:
                case PpSlideLayout.ppLayoutComparison:
                case PpSlideLayout.ppLayoutTwoObjects:
                case PpSlideLayout.ppLayoutObject: // Erstes Shape = Text -> Reguläre Folie, ansonsten "Normale Folie" ohne Titelüberschrift -> Verhalten wie bei Bilderfolie
                    Microsoft.Office.Interop.PowerPoint.Shape firstshape = presentation.Slides[pageNumber].Shapes[1]; // Erstes Shape -> Text
                    return (firstshape.HasTextFrame == MsoTriState.msoTrue) && (firstshape.TextFrame.HasText == MsoTriState.msoTrue)
                        ? SlideType.Slide : SlideType.ImageSlide;
                case PpSlideLayout.ppLayoutTitleOnly:
                case PpSlideLayout.ppLayoutBlank:
                case PpSlideLayout.ppLayoutCustom:
                case PpSlideLayout.ppLayoutTable:
                case PpSlideLayout.ppLayoutOrgchart:
                    return SlideType.ImageSlide;       // TODO: Tabellen mit eigenem LayoutType oder Blank ?   
                default:
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Unknown SlideType detected:{presentation.Slides[pageNumber].Layout} in page {pageNumber}");
                    return SlideType.Unknown;
            }
        }

        public (bool isLayoutValid, List<int> pagesWithInvalidLayout) IsLayoutValid()
        {
            List<int> listOfWrongPages = new List<int>();
            bool isValid = true;
            for (int pagenumber = 1; pagenumber <= presentation.Slides.Count; pagenumber++)
            {
                if (GetSlideType(pagenumber) == SlideType.Unknown)
                {
                    isValid = false;
                    listOfWrongPages.Add(pagenumber);
                }
            }
            return (isValid, listOfWrongPages);
        }
    }
}
