using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Reflection;
using ppedv.pocgen.Domain.Interfaces;

namespace ppedv.pocgen.Domain.Models
{
    public class PowerPointPresentationOpener : IOfficeFileOpener<IPowerPointPresentation>, IDisposable
    {
        private Application application;
        public PowerPointPresentationOpener(string[] ValidExtensions)
        {
            this.ValidExtensions = ValidExtensions ?? new string[] { ".ppt", ".pptx" };
            this.application = new Application();
        }
        public string[] ValidExtensions { get;}

        public void Dispose()
        {
            application?.Quit();
            application = null;
            MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"PowerPointPresentationOpener: disposed" ));
        }

        public IPowerPointPresentation OpenFile(string fileName)
        {
            if (!System.IO.File.Exists(fileName))
            {
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"PowerPointPresentationOpener OpenFile: File {fileName} not found" ));
                throw new System.IO.FileNotFoundException("Die angegebene PowerPoint-Datei wurde nicht gefunden.", fileName);
            }
            return new PowerPointPresentation(application.Presentations.Open(fileName, WithWindow: MsoTriState.msoFalse, ReadOnly: MsoTriState.msoTrue)); // TODO: System.IO.FileLoadException, wenn die datei von einem anderen prozess bereits genutzt wird (zb parallel offen in PP) - System.IO.FileLoadException occurred
        }
    }
}
