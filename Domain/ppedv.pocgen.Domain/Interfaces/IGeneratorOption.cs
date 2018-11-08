namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IGeneratorOption
    {
        string ID { get; }
        string Description { get; }
        bool IsEnabled { get; set; }
    }
}
