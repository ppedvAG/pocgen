using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;

namespace pocgen.Contracts.Interfaces
{
    public interface IPowerPointPresentation : IOfficeFile
    {
        Slides Slides { get; }
        int NumberOfSlidesInPresentation { get; }
        SlideType GetSlideType(int pageNumber);
        (bool isLayoutValid, List<int> pagesWithInvalidLayout) IsLayoutValid();
    }
}
