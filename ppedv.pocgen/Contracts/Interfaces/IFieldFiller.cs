using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;

namespace pocgen.Contracts.Interfaces
{
    public interface IFieldFiller
    {
        void FillFieldWithInfo(Field field, Slide correspondingSlide, ICourseInfo courseInfo);
    }
}
