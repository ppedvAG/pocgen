using pocgen.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace pocgen.Contracts.Models
{
    public class WordDocument : IWordDocument
    {
        private Document document;
        public WordDocument(Document document)
        {
            this.document = document;
        }

        public InlineShapes InlineShapes => document.InlineShapes;

        public Sections Sections => document.Sections;

        public Selection Selection => document.Application.Selection;

        public Range Content => document.Content;

        public void Dispose()
        {
            try
            {
                this.document?.Close(false);
                this.document = null;
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"WordDocument: disposed" ));
            }
            catch (Exception) // Liegt an Office, manchmal ist sogar ein "ForceClose" nicht möglich
            {
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Exception when trying to dispose: Close() not possible"));
            }
        }

        public Range Range(int start, int end) => document.Range(start, end);

        public void SetImageSyle() //TODO: Auslagern auf einen "ImageStyleSetter" ?
        {
            MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Start setting ImageStyle"));
            foreach (InlineShape shape in document.InlineShapes)
            {
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    shape.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;
                    shape.Borders.OutsideLineWidth = WdLineWidth.wdLineWidth100pt;
                    shape.Borders.OutsideColor = WdColor.wdColorAutomatic;
                }
            }
            MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"ImageStyle set successfully"));
        }
    }
}
