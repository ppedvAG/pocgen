using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Reflection;
using ppedv.pocgen.Domain.Interfaces;

namespace ppedv.pocgen.Domain.Models
{
    public class WordDocumentOpener : IOfficeFileOpener<WordDocument>
    {
        private readonly Application application;
        public WordDocumentOpener(Application application, string[] ValidExtensions)
        {
            this.application = application;
            this.ValidExtensions = ValidExtensions ?? new string[] { ".doc", ".docx" };
        }
        public string[] ValidExtensions { get; }

        public WordDocument OpenFile(string fileName)
        {
            if (!System.IO.File.Exists(fileName))
            {
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"File {fileName} not found" ));
                throw new System.IO.FileNotFoundException("Die angegebene Word-Datei wurde nicht gefunden.", fileName);
            }
            return new WordDocument(application.Documents.Open(fileName, Visible: false));
        }
    }
}
