using Microsoft.Office.Interop.Word;

namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IWordDocument : IOfficeFile
    {
        InlineShapes InlineShapes { get; }
        Sections Sections { get; }
        Selection Selection { get; }
        Range Content { get; }
        Range Range(int start, int end);
        void SetImageSyle();
    }
}
