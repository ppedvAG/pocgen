using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;
using Microsoft.Practices.Unity;
using pocgen.Contracts.Models;
using pocgen.Contracts.Interfaces;
using System.Windows.Forms.Integration;
using pocgen.View.Controls;
using pocgen.ViewModels.ViewModel;

namespace pocgen_WordAddIn
{
    public partial class pocgen_Ribbon
    {
        private Microsoft.Office.Tools.CustomTaskPane ctp;
        private bool isInit = false;

        private void pocgen_Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonOpen_Click(object sender, RibbonControlEventArgs e)
        {
            if (!isInit)
            {
                Globals.ThisAddIn.unityContainer = new UnityContainer();

                Globals.ThisAddIn.unityContainer.RegisterInstance<IOfficeFileOpener<IWordDocument>>(new WordDocumentOpener(Globals.ThisAddIn.Application, new string[] { ".doc", ".docx", ".dot", ".dotx" }));
                Globals.ThisAddIn.unityContainer.RegisterInstance<IOfficeFileOpener<IPowerPointPresentation>>(new PowerPointPresentationOpener(new string[] { ".ppt", ".pptx" }));
                Globals.ThisAddIn.unityContainer.RegisterInstance<IMarkdownGenerator>(new MarkdownGenerator());
                Globals.ThisAddIn.unityContainer.RegisterType<IGenerator, Generator>();
                Globals.ThisAddIn.unityContainer.RegisterType<IFieldFiller, FieldFiller>();
                Globals.ThisAddIn.unityContainer.RegisterInstance<IWordDocument>("ActiveDocument", new WordDocument(Globals.ThisAddIn.Application.ActiveDocument));

                var usercontrol = new System.Windows.Forms.UserControl();
                var elementhost = new ElementHost();
                var vm = new WordViewModel(
                    Globals.ThisAddIn.unityContainer.Resolve<IOfficeFileOpener<IWordDocument>>(),
                    Globals.ThisAddIn.unityContainer.Resolve<IOfficeFileOpener<IPowerPointPresentation>>(),
                    Globals.ThisAddIn.unityContainer.Resolve<IGenerator>(),
                    new List<IGeneratorOption>
                    {
                        new GeneratorOption("ISBeakAtStart","Seitenumbruch beim Anfang einer Reihe von Bilderfolien"),
                        new GeneratorOption("ISBreakBetween","Seitenumbrich zwischen einzelnen Bilderfolien"),
                    },
                    Globals.ThisAddIn.unityContainer.Resolve<IWordDocument>("ActiveDocument"));
                elementhost.Child = new MainViewControl(vm);

                elementhost.Dock = System.Windows.Forms.DockStyle.Fill;
                usercontrol.Controls.Add(elementhost);

                ctp = Globals.ThisAddIn.CustomTaskPanes.Add(usercontrol, "poc - Generator");
                ctp.Width = 620;
                isInit = true;
                ctp.VisibleChanged += (s,ea) =>
                {
                    vm.Cleanup();
                };
            }
            ctp.Visible = !ctp.Visible;
        }
    }
}
