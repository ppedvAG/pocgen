using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Interfaces
{
    public interface IGeneratorOption
    {
        string ID { get; }
        string Description { get; }
        bool IsEnabled { get; set; }
    }
}
