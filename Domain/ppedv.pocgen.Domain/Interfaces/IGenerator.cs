using System;
using System.Collections.Generic;
namespace ppedv.pocgen.Domain.Interfaces
{
    public interface IGenerator
    {
        event EventHandler<IGeneratorEventArgs> GeneratorProgressChanged;
        void GenerateDocument(IEnumerable<string> usedPowerPointPresentations, IWordDocument templateForOutputDocument, IWordDocument outputDocument, ICollection<IGeneratorOption> generatorOptions);
    }
}
