using System.IO;

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
        public string PreviewImagePath { get; set; }
        public (int,int) PreviewImageRange { get; set; }

        private bool isIncluded;
        public bool IsIncluded
        {
            get => isIncluded;
            set => SetValue(ref isIncluded, value);
        }
    }
}
