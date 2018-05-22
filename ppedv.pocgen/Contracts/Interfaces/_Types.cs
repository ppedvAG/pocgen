using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Interfaces
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
