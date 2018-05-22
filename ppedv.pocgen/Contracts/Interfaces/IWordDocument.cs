using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;

namespace pocgen.Contracts.Interfaces
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
