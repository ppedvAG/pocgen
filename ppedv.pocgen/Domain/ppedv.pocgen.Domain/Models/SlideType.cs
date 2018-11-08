using System;
using System.Collections.Generic;
using System.Text;

namespace ppedv.pocgen.Domain.Models
{
    public enum SlideType
    {
        None,               // Startwert -> Erste Seite braucht keinen Zeilenumbruch -> Im Switch passiert nix
        Title,              // Titelfolie
        Slide,              // reguläre Folie
        ImageSlide,         // Bild/Screenshotfolie
        Unknown             // Unbekannt -> Muss im Programm nachgetragen werden !
    }
}
