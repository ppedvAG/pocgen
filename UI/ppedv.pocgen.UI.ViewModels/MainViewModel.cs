using System;
using System.Text;
using ppedv.pocgen.Domain.Interfaces;
using ppedv.pocgen.Domain.Models;
using ppedv.pocgen.Logic;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Input;

using Word = Microsoft.Office.Interop.Word;
using System.Threading.Tasks;

namespace ppedv.pocgen.UI.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        public MainViewModel()
        {
            this.wordFileOpener = new WordDocumentOpener(new string[] { ".doc", ".docx", ".dot", ".dotx" });
            this.powerPointFileOpener = new PowerPointPresentationOpener(new string[] { ".ppt", ".pptx" });
            this.GeneratorOptions = new ObservableCollection<IGeneratorOption>
            {
                new GeneratorOption("ISBeakAtStart","Seitenumbruch beim Anfang einer Reihe von Bilderfolien"),
                new GeneratorOption("ISBreakBetween","Seitenumbrich zwischen einzelnen Bilderfolien"),
            };
            this.generator = new Generator(powerPointFileOpener, new FieldFiller());

            PowerPointPresentations = new ObservableCollection<PowerPointPresentationItem>();
            IsValidFolderSelected = false;
            IsValidTemplateSelected = false;
            UIElementsEnabled = true;
            generator.GeneratorProgressChanged += (sender, e) => GeneratorProgressValue = e.TotalSlidesDone;
        }
        ~MainViewModel()
        {
            templateForOutputDocument?.Dispose();
            wordFileOpener?.Dispose();
            powerPointFileOpener?.Dispose();
        }

        private IWordDocument templateForOutputDocument;
        private readonly IGenerator generator;
        private readonly IOfficeFileOpener<IWordDocument> wordFileOpener;
        private readonly IOfficeFileOpener<IPowerPointPresentation> powerPointFileOpener;

        private string dateFilter;
        public string DateFilter
        {
            get => dateFilter;
            set => SetValue(ref dateFilter, value);
        }

        private string messageFilter;
        public string MessageFilter
        {
            get => messageFilter;
            set => SetValue(ref messageFilter, value);
        }

        private string classFilter;
        public string ClassFilter
        {
            get => classFilter;
            set => SetValue(ref classFilter, value);
        }

        private string memberFilter;
        public string MemberFilter
        {
            get => memberFilter;
            set => SetValue(ref memberFilter, value);
        }

        private string templatePath;
        public string TemplatePath
        {
            get => templatePath;
            set
            {
                if (string.IsNullOrWhiteSpace(value)) // Reset
                {
                    SetValue(ref templatePath, string.Empty);
                    IsValidTemplateSelected = false;
                    return;
                }
                if (SetValue(ref templatePath, value))
                {
                    templateForOutputDocument?.Dispose(); // Falls schon eins offen sein sollte und ein neues gewählt wird
                    if (wordFileOpener.ValidExtensions.Contains(Path.GetExtension(value)) && File.Exists(value))
                    {
                        templateForOutputDocument = wordFileOpener.OpenFile(value);
                        IsValidTemplateSelected = true;
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Valid TemplatePath selected");
                    }
                    else
                    {
                        IsValidTemplateSelected = false;
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Invalid TemplatePath selected");
                    }
                }
            }
        }

        private bool isValidTemplateSelected;
        public bool IsValidTemplateSelected
        {
            get => isValidTemplateSelected;
            set => SetValue(ref isValidTemplateSelected, value);
        }

        private string folderPath;
        public string FolderPath
        {
            get => folderPath;
            set
            {
                if (string.IsNullOrWhiteSpace(value) || !Directory.Exists(value)) // Reset
                {
                    SetValue(ref folderPath, string.Empty);
                    IsValidFolderSelected = false;
                    return;
                }
                if (SetValue(ref folderPath, value))
                {
                    PowerPointPresentations.Clear();

                    void GetPowerPointPresentationPathsFromDirectory(string directory)
                    {
                        foreach (string subdirectory in Directory.GetDirectories(directory))
                        {
                            GetPowerPointPresentationPathsFromDirectory(subdirectory);
                        }
                        foreach (string file in Directory.GetFiles(directory))
                        {
                            if (powerPointFileOpener.ValidExtensions.Contains(Path.GetExtension(file)))
                            {
                                PowerPointPresentationItem pppi = new PowerPointPresentationItem(powerPointFileOpener) { FileName = file, IsIncluded = false };
                                pppi.PropertyChanged += (sender, e) =>
                                {
                                    if (e.PropertyName.Equals(nameof(PowerPointPresentationItem.IsIncluded)))
                                    {
                                        GeneratorProgressMaximum = PowerPointPresentations.Where(x => x.IsIncluded == true)
                                                                                          .Sum(x => x.NumberOfSlidesInPresentation);
                                    }
                                };
                                PowerPointPresentations.Add(pppi);
                            }
                        }
                    }
                    GetPowerPointPresentationPathsFromDirectory(folderPath);

                    if (PowerPointPresentations.Count > 0)
                    {
                        IsValidFolderSelected = true;
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Valid FolderPath selected");
                    }
                    else
                    {
                        IsValidFolderSelected = false;
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Invalid FolderPath selected");
                    }
                }
            }
        }

        private bool isValidFolderSelected;
        public bool IsValidFolderSelected
        {
            get => isValidFolderSelected;
            set => SetValue(ref isValidFolderSelected, value);
        }

        private int generatorProgressMaximum;
        public int GeneratorProgressMaximum
        {
            get => generatorProgressMaximum;
            set => SetValue(ref generatorProgressMaximum, value);
        }

        private int generatorProgressValue;
        public int GeneratorProgressValue
        {
            get => generatorProgressValue;
            set => SetValue(ref generatorProgressValue, value);
        }

        private bool uiElementsEnabled;
        public bool UIElementsEnabled
        {
            get => uiElementsEnabled;
            set => SetValue(ref uiElementsEnabled, value);
        }

        public ICollection<PowerPointPresentationItem> PowerPointPresentations { get; set; }
        public ICollection<IGeneratorOption> GeneratorOptions { get; set; }

        private ICommand buttonSelectTemplateClickCommand;
        public ICommand ButtonSelectTemplateClickCommand
        {
            get
            {
                buttonSelectTemplateClickCommand = buttonSelectTemplateClickCommand ?? new RelayCommand(parameter =>
                {
                    OpenFileDialog dlg = new OpenFileDialog();
                    if (dlg.ShowDialog() == DialogResult.OK)
                        TemplatePath = dlg.FileName;
                });
                return buttonSelectTemplateClickCommand;
            }
        }

        private ICommand buttonSelectFolderClickCommand;
        public ICommand ButtonSelectFolderClickCommand
        {
            get
            {
                buttonSelectFolderClickCommand = buttonSelectFolderClickCommand ?? new RelayCommand(parameter =>
                {
                    FolderBrowserDialog dlg = new FolderBrowserDialog();
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        FolderPath = dlg.SelectedPath;
                        GeneratorProgressMaximum = 0;
                    }
                });
                return buttonSelectFolderClickCommand;
            }
        }

        private ICommand buttonStartClickCommand;
        public ICommand ButtonStartClickCommand
        {
            get
            {
                buttonStartClickCommand = buttonStartClickCommand ?? new RelayCommand(parameter =>
                {
                    Task.Run(() =>
                    {
                        var wordInstance = new Word.Application();
                        var outputDocument = new WordDocument(wordInstance.Documents.Add());
                        UIElementsEnabled = false;
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Generator-Start");
                        generator.GenerateDocument(PowerPointPresentations
                                .Where(x => x.IsIncluded)
                                .Select(x => x.FileName), templateForOutputDocument, outputDocument, GeneratorOptions);
                        Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Generator-Finish");
                        UIElementsEnabled = true;
                        wordInstance.Visible = true;
                    });
                });
                return buttonStartClickCommand;
            }
        }

        private ICommand buttonResetClickCommand;
        public ICommand ButtonResetClickCommand
        {
            get
            {
                buttonResetClickCommand = buttonResetClickCommand ?? new RelayCommand(parameter =>
                {
                    UIElementsEnabled = true;
                    TemplatePath = string.Empty;
                    FolderPath = string.Empty;

                    IsValidTemplateSelected = false;
                    IsValidFolderSelected = false;

                    PowerPointPresentations.Clear();
                    GeneratorProgressMaximum = 0;
                    GeneratorProgressValue = 0;

                    foreach (IGeneratorOption option in GeneratorOptions)
                        option.IsEnabled = false;

                    templateForOutputDocument?.Dispose();
                });
                return buttonResetClickCommand;
            }
        }

        private ICommand buttonSelectAllPresentationsClickCommand;
        public ICommand ButtonSelectAllPresentationsClickCommand
        {
            get
            {
                buttonSelectAllPresentationsClickCommand = buttonSelectAllPresentationsClickCommand ?? new RelayCommand(parameter =>
                {
                    foreach (PowerPointPresentationItem item in PowerPointPresentations)
                        item.IsIncluded = true;
                });
                return buttonSelectAllPresentationsClickCommand;
            }
        }

        private ICommand buttonUnselectAllPresentationsClickCommand;
        public ICommand ButtonUnselectAllPresentationsClickCommand
        {
            get
            {
                buttonUnselectAllPresentationsClickCommand = buttonUnselectAllPresentationsClickCommand ?? new RelayCommand(parameter =>
                {
                    foreach (PowerPointPresentationItem item in PowerPointPresentations)
                        item.IsIncluded = false;
                });
                return buttonUnselectAllPresentationsClickCommand;
            }
        }
        public void Cleanup()
        {
            ButtonResetClickCommand?.Execute(null);
        }
    }
}
