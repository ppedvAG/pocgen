using ppedv.pocgen.Domain.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppedv.pocgen.Domain.Models
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
