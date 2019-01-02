using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Reflection;
using ppedv.pocgen.Domain.Interfaces;
using ppedv.pocgen.Domain.Models;
using System.Diagnostics;

namespace ppedv.pocgen.Logic
{
    public class Generator : IGenerator
    {
        public Generator(IOfficeFileOpener<IPowerPointPresentation> fileOpener, IFieldFiller fieldFiller)
        {
            this.fileOpener = fileOpener;
            this.fieldFiller = fieldFiller;

            courseInfo = new CourseInfo();
        }

        private ICourseInfo courseInfo;
        private IOfficeFileOpener<IPowerPointPresentation> fileOpener;
        private IFieldFiller fieldFiller;
        public event EventHandler<IGeneratorEventArgs> GeneratorProgressChanged;

        public void GenerateDocument(IEnumerable<string> usedPowerPointPresentations, IWordDocument templateForOutputDocument, IWordDocument outputDocument, ICollection<IGeneratorOption> generatorOptions)
        {
            bool isFirstModule = true;
            int totalSlidesDone = 0;
            GeneratorProgressChanged?.Invoke(this, new GeneratorEventArgs(totalSlidesDone));

            foreach (string pathToPowerPointPresentation in usedPowerPointPresentations)
            {
                var presentation = fileOpener.OpenFile(pathToPowerPointPresentation);
                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Opened file {pathToPowerPointPresentation}");

                if (!isFirstModule)
                {
                    InsertNewSectionIntoOutputDocument(outputDocument);
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] New section inserted");
                }

                for (int currentSlideNumber = 1; currentSlideNumber <= presentation.Slides.Count; currentSlideNumber++)
                {
                    JumpToLastPositionInDocumentAndSetCursor(outputDocument);

                    int outputDocumentStartOfCurentPage = outputDocument.Selection.Start;

                    // Dieser Switch schaut nach, was für ein interner "Folientyp" die aktuelle Folie ist und wird basierend auf der letzten Folie Seitenumbrüche oder neue Sektionen einfüge
                    switch (presentation.GetSlideType(currentSlideNumber))
                    {
                        case SlideType.Title:
                            #region Kurs und Modulinformationen für den Header zwischenspeichern
                            if (isFirstModule && currentSlideNumber == 1) // Modul00 - Titeltext
                            {
                                courseInfo.CourseName = GetTitleTextFromSlideInPresentation(presentation,currentSlideNumber).Replace('–', '-').Replace('—', '-'); // Powerpoint macht aus einem Bindestrich oftmals ein En-Dash oder Em-Dash -> Vergleichsfehler
                                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] got courseInfo from first slide in {pathToPowerPointPresentation}");
                            }
                            else // ModulXX - Titeltext
                            {
                                if(currentSlideNumber != 1) // Neue pptx -> keine doppelter Seitenumbruch
                                    outputDocument.Range(outputDocumentStartOfCurentPage, outputDocumentStartOfCurentPage).InsertBreak(WdBreakType.wdPageBreak);
                                JumpToLastPositionInDocumentAndSetCursor(outputDocument);
                                courseInfo.CourseCurrentModuleName = GetTitleTextFromSlideInPresentation(presentation,currentSlideNumber).Replace('–','-').Replace('—','-'); // Powerpoint macht aus einem Bindestrich oftmals ein En-Dash oder Em-Dash -> Vergleichsfehler
                                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] got current module name from first slide in {pathToPowerPointPresentation}");
                            }
                            #endregion
                            CopyTemplateToClipboardAndPasteIntoOutputDocument(templateForOutputDocument, outputDocument);
                            break;
                        case SlideType.Slide:
                            outputDocument.Range(outputDocumentStartOfCurentPage, outputDocumentStartOfCurentPage).InsertBreak(WdBreakType.wdPageBreak);
                            JumpToLastPositionInDocumentAndSetCursor(outputDocument);
                            CopyTemplateToClipboardAndPasteIntoOutputDocument(templateForOutputDocument, outputDocument);
                            break;
                        case SlideType.ImageSlide:
                            SlideType lastSlideType = (currentSlideNumber == 1) ? SlideType.None : presentation.GetSlideType(currentSlideNumber - 1);
                            // Wenn eine neue Folie ohne Titel kommt => Entscheidung bez. Standardverhalten (PageBreak am Anfang)
                            if (lastSlideType != SlideType.ImageSlide && generatorOptions.First(x => x.ID == "ISBeakAtStart").IsEnabled ||
                               // Wenn mehrere Folien ohne Titel kommen => Entscheidung bez. Standardverhalten (PageBreak dazwischen)
                               lastSlideType == SlideType.ImageSlide && generatorOptions.First(x => x.ID == "ISBreakBetween").IsEnabled)
                            {
                                outputDocument.Range(outputDocumentStartOfCurentPage, outputDocumentStartOfCurentPage).InsertBreak(WdBreakType.wdPageBreak);
                                JumpToLastPositionInDocumentAndSetCursor(outputDocument);
                                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] inserted break for ImageSlide according to option set in UI");
                            }

                            CopyTemplateToClipboardAndPasteIntoOutputDocument(templateForOutputDocument, outputDocument);
                            JumpToLastPositionInDocumentAndSetCursor(outputDocument);

                            foreach (Field field in outputDocument.Range(outputDocumentStartOfCurentPage, outputDocument.Selection.End).Fields)
                            {
                                if (field.Code.Text.Trim().Equals("Inhalt"))
                                {
                                    field.Unlink();
                                    break;
                                }
                            }

                            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] removed Inhalt-Tag for screenshot-page {currentSlideNumber}");

                            break;
                    }

                    JumpToLastPositionInDocumentAndSetCursor(outputDocument);
                    int outputDocumenEndOfCurrentPage = outputDocument.Selection.End;

                    FillAllFieldsForTheCurrentPage(outputDocument, presentation, currentSlideNumber, outputDocumentStartOfCurentPage, outputDocumenEndOfCurrentPage);

                    GeneratorProgressChanged?.Invoke(this, new GeneratorEventArgs(++totalSlidesDone));
                }

                FillHeaderAndFooterForFinishedSection(outputDocument, presentation);

                if (isFirstModule)
                    isFirstModule = false;
            }
            outputDocument.SetImageSyle();
        }

        private void FillHeaderAndFooterForFinishedSection(IWordDocument outputDocument, IPowerPointPresentation presentation)
        {
            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Start inserting Headers for current section");

            foreach (Field field in outputDocument.Sections[outputDocument.Sections.Count].Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields)
                fieldFiller.FillFieldWithInfo(field, presentation.Slides[presentation.Slides.Count], courseInfo);

            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Start inserting Footers for current section");

            foreach (Field field in outputDocument.Sections[outputDocument.Sections.Count].Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range.Fields)
                fieldFiller.FillFieldWithInfo(field, presentation.Slides[presentation.Slides.Count], courseInfo);
        }

        private void FillAllFieldsForTheCurrentPage(IWordDocument outputDocument, IPowerPointPresentation presentation, int currentSlideNumber, int outputDocumentStartOfCurentPage, int outputDocumenEndOfCurrentPage)
        {
            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Start filling Fields for slide {currentSlideNumber}");

            foreach (Field field in outputDocument.Range(outputDocumentStartOfCurentPage, outputDocumenEndOfCurrentPage).Fields)
                fieldFiller.FillFieldWithInfo(field, presentation.Slides[currentSlideNumber], courseInfo);
            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] All Fields for slide {currentSlideNumber} filled");
        }

        private static string GetTitleTextFromSlideInPresentation(IPowerPointPresentation presentation,int slide)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in presentation.Slides[slide].Shapes)
            {
                if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    return shape.TextFrame.TextRange.Text;
                }
            }
            return string.Empty;
        }

        private static void JumpToLastPositionInDocumentAndSetCursor(IWordDocument outputDocument)
        {
            outputDocument.Selection.EndKey(WdUnits.wdStory, WdMovementType.wdExtend);
            outputDocument.Selection.Collapse(WdCollapseDirection.wdCollapseEnd);
        }

        private static void InsertNewSectionIntoOutputDocument(IWordDocument outputDocument)
        {
            Section s = outputDocument.Sections.Add();
            s.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
            s.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].LinkToPrevious = false;
        }

        private void CopyTemplateToClipboardAndPasteIntoOutputDocument(IWordDocument template, IWordDocument outputDocument)
        {
            int maxTries = 5;
            bool gotException;
            do
            {
                gotException = false;
                try
                {
                    Range templateRange = template.Range(template.Content.Start, template.Content.End);
                    templateRange.Copy();
                }
                catch (Exception) // Manchmal spinnt Word und braucht mehrere Versuche, ka warum
                {
                    gotException = true;
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Exception while trying to copy template into clipboard: left:{maxTries}");
                    if (--maxTries == 0)
                    {
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] maxTries reached while trying to copy template into clipboard");
                        break;
                    }
                }
            } while (gotException);
            outputDocument.Selection.Paste();
            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] template copied and pasted into outputDocument");
        }
    }
}
