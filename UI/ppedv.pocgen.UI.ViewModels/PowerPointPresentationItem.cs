using ppedv.pocgen.Domain.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ppedv.pocgen.UI.ViewModels
{
    public class PowerPointPresentationItem : BaseViewModel
    {
        public PowerPointPresentationItem(IOfficeFileOpener<IPowerPointPresentation> powerPointFileOpener)
        {
            this.powerPointFileOpener = powerPointFileOpener;
        }

        private readonly IOfficeFileOpener<IPowerPointPresentation> powerPointFileOpener;

        private string fileName;
        public string FileName
        {
            get => fileName;
            set => SetValue(ref fileName, value);
        }

        private bool isIncluded;
        public bool IsIncluded
        {
            get => isIncluded;
            set => SetValue(ref isIncluded, value);
        }

        private int numberOfSlidesInPresentation = -1;
        public int NumberOfSlidesInPresentation
        {
            get
            {
                if (numberOfSlidesInPresentation == -1)
                {
                    IPowerPointPresentation presentation = powerPointFileOpener.OpenFile(FileName);
                    numberOfSlidesInPresentation = presentation.NumberOfSlidesInPresentation;
                    presentation.Dispose();
                    presentation = null;
                }
                return numberOfSlidesInPresentation;
            }
        }

        public bool IsLayoutValid
        {
            get
            {
                IPowerPointPresentation presentation = powerPointFileOpener.OpenFile(FileName);
                var result = presentation.IsLayoutValid();
                presentation.Dispose();
                presentation = null;

                LayoutDescription = (result.isLayoutValid) ? null : $"Folgende Seiten haben ein ungültiges Layout: {string.Join(", ", result.pagesWithInvalidLayout)}";
                return result.isLayoutValid;
            }
        }

        private string layoutDescription;

        public string LayoutDescription
        {
            get { return layoutDescription; }
            set { SetValue(ref layoutDescription, value); }
        }

    }
}
