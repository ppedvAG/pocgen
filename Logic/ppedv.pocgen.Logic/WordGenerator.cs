using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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

namespace ppedv.pocgen.Logic
{
    public class WordGenerator
    {
        public void GeneratePOC_Document(string inputPresentationFullPath,string slideImageDirectoryPath, string outputPOCFullPath)
        {
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(outputPOCFullPath, WordprocessingDocumentType.Document))
            {
                #region Init Document
                wordDocument.AddMainDocumentPart();
                wordDocument.MainDocumentPart.Document = new Document();
                var body = wordDocument.MainDocumentPart.Document.AppendChild(new Body());
                InitStylesFor(wordDocument);
                #endregion

                using (PresentationDocument presentationDocument = PresentationDocument.Open(inputPresentationFullPath, false))
                {
                    for (int currentSlide = 0; currentSlide < presentationDocument.PresentationPart.SlideParts.Count(); currentSlide++)
                    {
                        string[] all = GetAllTextFromSlidePart(GetSlidePart(presentationDocument, currentSlide));
                        if (all == null)
                            continue;

                        bool firstElement = true;
                        foreach (var powerP in all)
                        {
                            if (firstElement)
                            {
                                body.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(powerP)))
                                {
                                    ParagraphProperties = new ParagraphProperties(new ParagraphStyleId() { Val = "Heading1" })
                                });
                                #region Bild einfügen
                                ImagePart imagePart = wordDocument.MainDocumentPart.AddImagePart(ImagePartType.Jpeg);
                                using (FileStream stream = new FileStream(Path.Combine(slideImageDirectoryPath, $"{currentSlide.ToString()}.png"), FileMode.Open))
                                {
                                    imagePart.FeedData(stream);
                                    using (Bitmap img = new Bitmap(stream))
                                    {
                                        const int maxWidthCm = 15;
                                        const int emusPerInch = 914400;
                                        const int emusPerCm = 360000;
                                        var widthEmus = (long)(img.Width / img.HorizontalResolution * emusPerInch);
                                        var heightEmus = (long)(img.Height / img.VerticalResolution * emusPerInch);
                                        var maxWidthEmus = (long)(maxWidthCm * emusPerCm);
                                        if (widthEmus > maxWidthEmus) // Wenn das Bild zu groß ist, runterskalieren
                                        {
                                            var ratio = (heightEmus * 1.0m) / widthEmus;
                                            widthEmus = maxWidthEmus;
                                            heightEmus = (long)(widthEmus * ratio);
                                        }
                                        AddImageToBody(wordDocument, wordDocument.MainDocumentPart.GetIdOfPart(imagePart), widthEmus, heightEmus);
                                    }
                                }
                                #endregion
                                firstElement = false;
                            }
                            else
                                body.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(powerP))));
                        }
                        string notes = GetAllNotesFromSlidePart(GetSlidePart(presentationDocument, currentSlide));
                        if (notes != null)
                            body.AppendChild(new Paragraph(new Run(new DocumentFormat.OpenXml.Wordprocessing.Text(notes))));

                        if (currentSlide != presentationDocument.PresentationPart.SlideParts.Count() - 1)
                            body.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
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

        private SlidePart GetSlidePart(PresentationDocument presentationDocument, int slideIndex)
        {
            if (slideIndex < 0)
                throw new ArgumentOutOfRangeException("slideIndex");

            if (presentationDocument?.PresentationPart?.Presentation?.SlideIdList != null)
            {
                // Get the collection of slide IDs from the slide ID list.
                var slideIds = presentationDocument.PresentationPart.Presentation.SlideIdList.ChildElements;
                // If the slide ID is in range...
                if (slideIndex < slideIds.Count)
                {
                    // Get the relationship ID of the slide.
                    string slidePartRelationshipId = (slideIds[slideIndex] as DocumentFormat.OpenXml.Presentation.SlideId).RelationshipId;
                    // Return the specified slide part from the relationship ID.
                    return (SlidePart)presentationDocument.PresentationPart.GetPartById(slidePartRelationshipId);
                }
            }
            return null;
        }
        private string[] GetAllTextFromSlidePart(SlidePart slidePart)
        {
            // Create a new linked list of strings.
            LinkedList<string> texts = new LinkedList<string>();

            // If the slide exists...
            if (slidePart?.Slide != null)
            {
                // Iterate through all the paragraphs in the slide.
                foreach (var paragraph in slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    // Create a new string builder.                    
                    StringBuilder paragraphText = new StringBuilder();
                    // Iterate through the lines of the paragraph.
                    foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        // Append each line to the previous lines.
                        paragraphText.Append(text.Text);
                    }

                    if (paragraphText.Length > 0)
                    {
                        // Add each paragraph to the linked list.
                        texts.AddLast(paragraph.InnerText);
                    }
                }
            }
            if (texts.Count > 0)
                return texts.ToArray();

            return null;
        }
        private string GetAllNotesFromSlidePart(SlidePart slidePart)
        {
            if (slidePart.NotesSlidePart != null && !string.IsNullOrWhiteSpace(slidePart.NotesSlidePart.NotesSlide.InnerText))
            // return slidePart.NotesSlidePart.NotesSlide.InnerText;
            {
                StringBuilder paragraphText = new StringBuilder();
                foreach (var paragraph in slidePart.NotesSlidePart.NotesSlide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>())
                {
                    // Create a new string builder.                    
                    // Iterate through the lines of the paragraph.
                    foreach (var text in paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>())
                    {
                        if (paragraph.InnerXml.Contains(" type=\"slidenum\" "))
                            continue;
                        // Append each line to the previous lines.
                        paragraphText.AppendLine(text.Text);
                    }
                }
                if (paragraphText.Length > 0)
                    return paragraphText.ToString();
            }
            return null;
        }
        private void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, long iWidth, long iHeight)
        {
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = iWidth, Cy = iHeight },
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
                                                new A.Extents() { Cx = iWidth, Cy = iHeight }),
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

            wordDoc.MainDocumentPart.Document.Body.AppendChild(p);
        }
    }
}
