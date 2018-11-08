using pocgen.Contracts.Interfaces;
using pocgen.Contracts.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Unity;

using PowerPoint =  Microsoft.Office.Interop.PowerPoint;
using Word =  Microsoft.Office.Interop.Word;
using pocgen.ViewModels.ViewModel;

namespace pocgen_WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            unityContainer = new UnityContainer();

            appInstance = new Word.Application();
            unityContainer.RegisterInstance<IOfficeFileOpener<IWordDocument>>(new WordDocumentOpener(appInstance, new string[] { ".doc", ".docx", ".dot", ".dotx" }));
            unityContainer.RegisterInstance<IOfficeFileOpener<IPowerPointPresentation>>(new PowerPointPresentationOpener(new string[] { ".ppt", ".pptx" }));
            unityContainer.RegisterInstance<IMarkdownGenerator>(new MarkdownGenerator());
            unityContainer.RegisterType<IGenerator, Generator>();
            unityContainer.RegisterType<IFieldFiller, FieldFiller>();

            DataContext =  new WPFViewModel(
                unityContainer.Resolve<IOfficeFileOpener<IWordDocument>>(),
                unityContainer.Resolve<IOfficeFileOpener<IPowerPointPresentation>>(),
                unityContainer.Resolve<IGenerator>(),
                appInstance,
                new List<IGeneratorOption>
                {
                        new GeneratorOption("ISBeakAtStart","Seitenumbruch beim Anfang einer Reihe von Bilderfolien"),
                        new GeneratorOption("ISBreakBetween","Seitenumbrich zwischen einzelnen Bilderfolien"),
                });
            InitializeComponent();
        }
        private Word.Application appInstance;
        public UnityContainer unityContainer { get; set; }

    }
}
