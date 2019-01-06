using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppedv.pocgen.UI.WPF.ViewModels
{
    public class PowerPointPresentationItem : BaseViewModel
    {
        public PowerPointPresentationItem(string fullPath)
        {
            FullPath = fullPath;
            FileName = Path.GetFileName(FullPath);
        }
        public string FullPath { get; set; }
        public string FileName { get; set; }

        private bool isIncluded;
        public bool IsIncluded
        {
            get => isIncluded;
            set => SetValue(ref isIncluded, value);
        }
    }
}
