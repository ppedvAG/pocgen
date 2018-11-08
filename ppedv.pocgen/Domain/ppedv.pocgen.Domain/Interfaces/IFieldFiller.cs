using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.PowerPoint;

namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IFieldFiller
    {
        void FillFieldWithInfo(Field field, Slide correspondingSlide, ICourseInfo courseInfo);
    }
}
