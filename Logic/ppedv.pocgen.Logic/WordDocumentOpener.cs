using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Reflection;
using ppedv.pocgen.Domain.Interfaces;
using ppedv.pocgen.Domain.Models;
using System.Diagnostics;

namespace ppedv.pocgen.Logic
{
    public class WordDocumentOpener : IOfficeFileOpener<WordDocument>
    {
        public WordDocumentOpener(string[] ValidExtensions)
        {
            this.application = new Application();
            this.ValidExtensions = ValidExtensions ?? new string[] { ".doc", ".docx" };
        }
        private Application application;
        public string[] ValidExtensions { get; }

        public WordDocument OpenFile(string fileName)
        {
            if (!System.IO.File.Exists(fileName))
            {
                Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] File {fileName} not found" );
                throw new System.IO.FileNotFoundException("Die angegebene Word-Datei wurde nicht gefunden.", fileName);
            }
            return new WordDocument(application.Documents.Open(fileName, Visible: false));
        }

        public void Dispose()
        {
            application?.Quit();
            application = null;
            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PowerPointPresentationOpener: disposed");
        }
    }
}
