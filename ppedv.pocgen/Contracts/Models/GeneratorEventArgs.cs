using pocgen.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Models
{
    public class GeneratorEventArgs : IGeneratorEventArgs
    {
        public GeneratorEventArgs(int TotalSlidesDone)
        {
            this.TotalSlidesDone = TotalSlidesDone;
        }
        public int TotalSlidesDone { get; }
    }
}
