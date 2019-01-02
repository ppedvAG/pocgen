﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Reflection;
using ppedv.pocgen.Domain.Interfaces;
using Microsoft.Office.Core;
using ppedv.pocgen.Domain.Models;
using System.Diagnostics;

namespace ppedv.pocgen.Logic
{
    public class FieldFiller : IFieldFiller
    {
        public void FillFieldWithInfo(Field field, Slide correspondingSlide, ICourseInfo courseInfo)
        {
            // Legende:
            // f.Code   == Feldcode ( zb "Page" -> Aktuelle Seitennummer)
            // f.Result == Ergebnis des berechneten Feldes
            // Bestimmte Texte im Code, die keinem richtigen Feldcode entsprechen, werden in dem Switch für meinen eigenen Code verwendet
            // f.Unlink(); zeigt letztendlich nur noch den Text an und zerstört das Feld
            string fieldName = field.Code.Text.Trim();
            switch (fieldName)
            {
                case "Überschrift":
                    string title = (correspondingSlide.Shapes.HasTitle == MsoTriState.msoTrue) ? correspondingSlide.Shapes.Title.TextFrame.TextRange.Text : string.Empty;
                    field.Result.Text = string.IsNullOrWhiteSpace(title) ? string.Empty : $"{correspondingSlide.Shapes.Title.TextEffect.Text}";
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}'");
                    field.Unlink();
                    break;
                case "Inhalt":
                    string slidetitle = (correspondingSlide.Shapes.HasTitle == MsoTriState.msoTrue) ? correspondingSlide.Shapes.Title.TextFrame.TextRange.Text : string.Empty;
                    for (int i = 1; i <= correspondingSlide.Shapes.Count; i++) //TODO: Inhalt-VERKEHRTHERUM-Fehler: falls der wieder kommen sollte nach meinem Refactoring, hier nachschauen !
                    {
                        if (correspondingSlide.Shapes[i].HasTextFrame == MsoTriState.msoTrue &&
                            correspondingSlide.Shapes[i].TextFrame.HasText == MsoTriState.msoTrue &&
                            !correspondingSlide.Shapes[i].TextFrame.TextRange.Text.Equals(slidetitle,StringComparison.InvariantCultureIgnoreCase)  &&
                            !correspondingSlide.Shapes[i].TextFrame.TextRange.Text.Replace('–', '-').Replace('—', '-').Equals(courseInfo.CourseCurrentModuleName,StringComparison.InvariantCultureIgnoreCase) &&
                            !Regex.IsMatch(correspondingSlide.Shapes[i].TextFrame.TextRange.Text, @"^\d+$"))
                        {
                            int maxtries = 3;
                            bool gotAnException = false;
                            do
                            {
                                try
                                {
                                    correspondingSlide.Shapes[i].TextFrame.TextRange.Copy();
                                    field.Application.Selection.PasteAndFormat(WdRecoveryType.wdFormatOriginalFormatting);
                                    gotAnException = false;
                                }
                                catch (Exception)
                                {
                                    gotAnException = true;
                                    if (--maxtries == 0)
                                    {
                                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] ERROR: Field '{fieldName}' could not paste Content");
                                        field.Result.Text = string.Empty; // Behebt "Fehler ! Textmarke nicht definiert."
                                        field.Unlink();
                                        return;
                                    }
                                }
                            } while (gotAnException);
                        }
                    }
                    field.Result.Text = string.Empty; // Behebt "Fehler ! Textmarke nicht definiert."
                    field.Unlink();
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}' with content from slide");
                    break;
                case "Notiz":
                    string notesInSlide = string.Empty;
                    if (correspondingSlide.NotesPage.Shapes?.Count >= 3) // Die Notizen sind immer im NotesPage.Shapes[2] drinnen !
                    {
                        try
                        {
                            notesInSlide = correspondingSlide.NotesPage.Shapes[2].TextFrame.TextRange.Text;
                        }
                        catch (ArgumentException)
                        { // TextFrame.TextRange kann in einigen lustigen Kombinationen in PowerPoint eine ArgumentException auslösen -> ignorieren
                            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Field: '{fieldName}' - no notes detected -> Exception ignored");
                        }
                        catch (Exception ex)
                        {
                            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Field: '{fieldName}' - Unknown Exception:{ex.Message}");
                        }
                    }
                    if (!string.IsNullOrWhiteSpace(notesInSlide)) // Wenn Notizen vorhanden sind -> Notizen ausgeben
                    {
                        field.Result.Text = notesInSlide; // -> Notizen 1:1 in Word übertragen
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}' with notes");
                    }
                    else
                        goto case "Inhalt";
                    field.Unlink();
                    break;
                case "Kursname":
                    field.Result.Text = courseInfo.CourseName;
                    field.Unlink();
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}'");
                    break;
                case "Modul":
                    field.Result.Text = courseInfo.CourseCurrentModuleName;
                    field.Unlink();
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}'");
                    break;
                case "Copyright":
                    field.Result.Text = "ppedv AG";
                    field.Unlink();
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}'");
                    break;
                case "Seite":
                    field.Code.Text = " Page"; // Page ist Feldfunktion, daher kein Unlink !
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully changed Field '{fieldName}' to '{field.Code.Text}'");
                    break;
                case "Slide":
                    int maxTries = 3;
                    bool gotException = false;
                    do
                    {
                        try
                        {
                            correspondingSlide.Copy();
                            field.Result.Paste();
                            gotException = false;
                        }
                        catch (Exception)
                        {
                            gotException = true;
                            if (--maxTries == 0)
                            {
                                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] ERROR: Field '{fieldName}' could not be filled with screenshot");
                                field.Unlink();
                                return;
                            }
                        }
                    } while (gotException);
                    field.Unlink();
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Successfully filled Field '{fieldName}'");
                    break;
                default:
                    field.Result.Text = $"--- Tag {fieldName} wurde nicht erkannt ---";
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Unknown Field: '{fieldName}'");
                    break;
            }
        }
    }
}
