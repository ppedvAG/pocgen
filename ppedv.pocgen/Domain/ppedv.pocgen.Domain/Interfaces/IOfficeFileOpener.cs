namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IOfficeFileOpener<out T> where T : IOfficeFile
    {
        T OpenFile(string fileName);
        string[] ValidExtensions { get; }
    }
}
