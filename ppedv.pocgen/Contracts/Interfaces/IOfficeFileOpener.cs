using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Interfaces
{
    public interface IOfficeFileOpener<out T> where T : IOfficeFile
    {
        T OpenFile(string fileName);
        string[] ValidExtensions { get; }
    }
}
