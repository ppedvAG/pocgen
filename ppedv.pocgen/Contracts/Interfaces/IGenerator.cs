using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Interfaces
{
    public interface IGenerator
    {
        event EventHandler<IGeneratorEventArgs> GeneratorProgressChanged;
        void GenerateDocument(IEnumerable<string> usedPowerPointPresentations, IWordDocument templateForOutputDocument, IWordDocument outputDocument, ICollection<IGeneratorOption> generatorOptions);
    }
}
