
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace ppedv.pocgen.Logic
{
    public class WordGenerator
    {
        public void GeneratePOC_Document(string inputPresentationFullPath, string slideImageDirectoryPath, string outputPOCFullPath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPOCFullPath, WordprocessingDocumentType.Document))
            {
                #region Init Document
                wordDocument.AddMainDocumentPart();
                wordDocument.MainDocumentPart.Document = new Document();
                var body = wordDocument.MainDocumentPart.Document.AppendChild(new Body());
                InitStylesFor(wordDocument);
                #endregion

                using (PresentationDocument inputPresentation = PresentationDocument.Open(inputPresentationFullPath, false))
                {
                    for (int currentSlide = 0; currentSlide < inputPresentation.PresentationPart.SlideParts.Count(); currentSlide++)
                    {
                        #region Read all Text from current slide
                        string[] allTextFromCurrentSlide = GetAllTextFromSlide(GetSlide(inputPresentation, currentSlide));
                        if (allTextFromCurrentSlide == null)
                            continue;
                        #endregion
                        InsertEachParagraphIntoWordDocument(wordDocument, allTextFromCurrentSlide, slideImageDirectoryPath, currentSlide);
                        InsertNotesIntoWordDocument(wordDocument, inputPresentation, currentSlide);
                        #region Insert PageBreak if Slide is not last
                        if (currentSlide != inputPresentation.PresentationPart.SlideParts.Count() - 1)
                            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                        #endregion
                    }
                }
            }
        }
        private void InitStylesFor(WordprocessingDocument wordDocument)
        {
            StyleDefinitionsPart styleDefinitions = wordDocument.MainDocumentPart.AddNewPart<StyleDefinitionsPart>();

            var styles = new Styles();
            styles.Save(styleDefinitions);
            styles = styleDefinitions.Styles;

            var style = new Style()
            {
                Type = StyleValues.Paragraph,
                StyleId = "Heading1",
                CustomStyle = true,
                Default = false
            };

            style.Append(new StyleName() { Val = "Heading 1" });

            var styleRunProperties = new StyleRunProperties();
            styleRunProperties.Append(new Bold());
            styleRunProperties.Append(new RunFonts() { Ascii = "Calibri" });
            styleRunProperties.Append(new FontSize() { Val = "40" });

            style.Append(styleRunProperties);

            styles.Append(style);
        }
        private void InsertEachParagraphIntoWordDocument(WordprocessingDocument wordDocument, string[] allTextFromCurrentSlide, string slideImageDirectoryPath, int currentSlide)
        {
            bool firstElement = true;
            foreach (var paragraph in allTextFromCurrentSlide)
            {
                if (firstElement)
                {
                    AppendParagraph(wordDocument, paragraph, "Heading1");
                    InsertImage(slideImageDirectoryPath, wordDocument, currentSlide);
                    firstElement = false;
                }
                else
                    AppendParagraph(wordDocument, paragraph);
            }
        }
        private void AppendParagraph(WordprocessingDocument wordDocument, string paragraph, string styleID = "")
        {
            wordDocument.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(new Text(paragraph))) { ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = styleID }) });
        }
        private void InsertNotesIntoWordDocument(WordprocessingDocument wordDocument, PresentationDocument inputPresentation, int currentSlide)
        {
            string notes = GetAllNotesFromSlide(GetSlide(inputPresentation, currentSlide));
            if (notes != null)
                AppendParagraph(wordDocument, notes);
        }
        private void InsertImage(string slideImageDirectoryPath, WordprocessingDocument wordDocument, int currentSlide)
        {
            ImagePart imagePart = wordDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
            using (FileStream stream = new FileStream(Path.Combine(slideImageDirectoryPath, $"{currentSlide.ToString()}.png"), FileMode.Open))
            {
                imagePart.FeedData(stream);
                using (Bitmap img = new Bitmap(stream))
                {
                    var resolution = GetImageResolutionInEMUs(img); // English Metric Units
                    var imageParagraph = CreateParagraphWithImage(wordDocument.MainDocumentPart.GetIdOfPart(imagePart), resolution.width, resolution.heigth);
                    wordDocument.MainDocumentPart.Document.Body.AppendChild(imageParagraph);
                }
            }
        }
        private (long width, long heigth) GetImageResolutionInEMUs(Bitmap img)
        {
            const int maxWidthCm = 15;
            const int emusPerInch = 914400;
            const int emusPerCm = 360000;
            long widthEmus = (long)(img.Width / img.HorizontalResolution * emusPerInch);
            long heightEmus = (long)(img.Height / img.VerticalResolution * emusPerInch);
            var maxWidthEmus = (long)(maxWidthCm * emusPerCm);
            if (widthEmus > maxWidthEmus) // Wenn das Bild zu groß ist, runterskalieren
            {
                var ratio = (heightEmus * 1.0m) / widthEmus;
                widthEmus = maxWidthEmus;
                heightEmus = (long)(widthEmus * ratio);
            }
            return (widthEmus, heightEmus);
        }
        private SlidePart GetSlide(PresentationDocument presentationDocument, int slideIndex)
        {
            if (presentationDocument?.PresentationPart?.Presentation?.SlideIdList != null)
            {
                var slideIds = presentationDocument.PresentationPart.Presentation.SlideIdList.ChildElements;
                if (slideIndex < slideIds.Count)
                {
                    string slidePartRelationshipId = (slideIds[slideIndex] as DocumentFormat.OpenXml.Presentation.SlideId).RelationshipId;
                    return (SlidePart)presentationDocument.PresentationPart.GetPartById(slidePartRelationshipId);
                }
            }
            return null;
        }
        private string[] GetAllTextFromSlide(SlidePart slidePart)
        {
            if (slidePart?.Slide != null)
            {
                LinkedList<string> texts = new LinkedList<string>();
                foreach (var paragraph in slidePart.Slide.Descendants<A.Paragraph>())
                {
                    StringBuilder paragraphText = new StringBuilder();
                    foreach (var text in paragraph.Descendants<A.Text>())
                        paragraphText.Append(text.Text);

                    if (paragraphText.Length > 0)
                        texts.AddLast(paragraph.InnerText);
                }
                if (texts.Count > 0)
                    return texts.ToArray();
            }
            return null;
        }
        private string GetAllNotesFromSlide(SlidePart slidePart)
        {
            if (slidePart.NotesSlidePart != null && !string.IsNullOrWhiteSpace(slidePart.NotesSlidePart.NotesSlide.InnerText))
            {
                StringBuilder paragraphText = new StringBuilder();
                foreach (var paragraph in slidePart.NotesSlidePart.NotesSlide.Descendants<A.Paragraph>())
                    foreach (var text in paragraph.Descendants<A.Text>())
                    {
                        if (paragraph.InnerXml.Contains(" type=\"slidenum\" "))
                            continue; // Zeilennummer ignorieren
                        paragraphText.AppendLine(text.Text);
                    }

                if (paragraphText.Length > 0)
                    return paragraphText.ToString();
            }
            return null;
        }
        private Paragraph CreateParagraphWithImage(string relationshipId, long width, long height)
        {
            #region Create Image-Element
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = width, Cy = height },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocProperties()
                         {
                             Id = 1U,
                             Name = "Picture 1"
                         },
                         new NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = 0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                             new A.Transform2D(
                                                new A.Offset() { X = 0L, Y = 0L },
                                                new A.Extents() { Cx = width, Cy = height }),
                                             new A.PresetGeometry(new A.AdjustValueList())
                                             { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U,
                     });
            #endregion
            #region Set Image - Style
            Paragraph p = new Paragraph(new Run(element))
            {
                ParagraphProperties = new ParagraphProperties()
                {
                    Justification = new Justification() { Val = JustificationValues.Center },
                    ParagraphBorders = new ParagraphBorders
                    {
                        TopBorder = new TopBorder() { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                        LeftBorder = new LeftBorder() { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                        BottomBorder = new BottomBorder() { Val = BorderValues.Thick, Size = 24, Color = "000000" },
                        RightBorder = new RightBorder() { Val = BorderValues.Thick, Size = 24, Color = "000000" }
                    }
                }
            };
            #endregion
            return p;
        }
    }
}