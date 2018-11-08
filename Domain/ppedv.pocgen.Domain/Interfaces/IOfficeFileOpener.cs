using System;

namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IOfficeFileOpener<out T> : IDisposable where T : IOfficeFile
    {
        T OpenFile(string fileName);
        string[] ValidExtensions { get; }
    }
}
