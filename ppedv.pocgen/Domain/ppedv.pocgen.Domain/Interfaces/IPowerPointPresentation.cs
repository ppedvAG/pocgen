using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using ppedv.pocgen.Domain.Models;

namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IPowerPointPresentation : IOfficeFile
    {
        Slides Slides { get; }
        int NumberOfSlidesInPresentation { get; }
        SlideType GetSlideType(int pageNumber);
        (bool isLayoutValid, List<int> pagesWithInvalidLayout) IsLayoutValid();
    }
}
