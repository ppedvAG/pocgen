using pocgen.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace pocgen.Contracts.Models
{
    public class GeneratorOption : IGeneratorOption
    {
        public GeneratorOption(string ID, string Description) : this(ID, Description, false) { }
        public GeneratorOption(string ID, string Description, bool IsEnabled)
        {
            this.ID = ID;
            this.Description = Description;
            this.IsEnabled = IsEnabled;
        }

        public string ID { get; }
        public string Description { get; }
        public bool IsEnabled { get; set; } // TODO: INotifyPropertyChanged für das UI ? BaseViewModel implementieren ?
    }
}
