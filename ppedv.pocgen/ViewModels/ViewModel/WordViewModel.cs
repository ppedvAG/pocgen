using pocgen.Contracts.Interfaces;
using pocgen.Contracts.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using System.Windows.Forms;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Reflection;

namespace pocgen.ViewModels.ViewModel
{
    public class WordViewModel : BaseViewModel
    {
        public WordViewModel(IOfficeFileOpener<IWordDocument> wordFileOpener, IOfficeFileOpener<IPowerPointPresentation> powerPointFileOpener, IGenerator generator, List<IGeneratorOption> GeneratorOptions, IWordDocument outputDocument)
        {
            this.wordFileOpener = wordFileOpener;
            this.powerPointFileOpener = powerPointFileOpener;
            this.GeneratorOptions = new ObservableCollection<IGeneratorOption>(GeneratorOptions);
            this.generator = generator;
            this.outputDocument = outputDocument;
            UILog = new ObservableCollection<LoggerEventArgs>();

            PowerPointPresentations = new ObservableCollection<PowerPointPresentationItem>();
            IsValidFolderSelected = false;
            IsValidTemplateSelected = false;
            UIElementsEnabled = true;
            generator.GeneratorProgressChanged += (sender, e) =>
            {
                GeneratorProgressValue = e.TotalSlidesDone;
            };

            generatorWorker = new BackgroundWorker();
            generatorWorker.DoWork += (sender, e) =>
            {
                UIElementsEnabled = false;
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Generator-Start"));
                generator.GenerateDocument(PowerPointPresentations
                        .Where(x => x.IsIncluded)
                        .Select(x => x.FileName), templateForOutputDocument, outputDocument, GeneratorOptions);
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Generator-Finish"));
                UIElementsEnabled = true;
            };

            MessagingCenter.Subscribe("Log", (object sender, EventArgs e) => DispatcherObject.Invoke(() => UILog.Add((e as LoggerEventArgs))));
        }

        private IWordDocument templateForOutputDocument;
        private IWordDocument outputDocument;
        private IGenerator generator;
        private IOfficeFileOpener<IWordDocument> wordFileOpener;
        private IOfficeFileOpener<IPowerPointPresentation> powerPointFileOpener;
        private readonly BackgroundWorker generatorWorker;

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
                        MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Valid TemplatePath selected"));

                    }
                    else
                    {
                        IsValidTemplateSelected = false;
                        MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Invalid TemplatePath selected"));
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
                        MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Valid FolderPath selected"));
                    }
                    else
                    {
                        IsValidFolderSelected = false;
                        MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Invalid FolderPath selected"));
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

        private ObservableCollection<LoggerEventArgs> uiLog;
        public ObservableCollection<LoggerEventArgs> UILog
        {
            get
            {
                IEnumerable<LoggerEventArgs> result = new List<LoggerEventArgs>();
                bool hadToFilter = false;
                if (!string.IsNullOrEmpty(DateFilter))
                {
                    result = result.Union(uiLog.Where(x => x.Time.ToLongTimeString().ToLower().Contains(DateFilter.ToLower())));
                    hadToFilter = true;
                }
                if (!string.IsNullOrEmpty(MessageFilter))
                {
                    result = result.Union(uiLog.Where(x => x.Message.ToLower().Contains(MessageFilter.ToLower())));
                    hadToFilter = true;
                }
                if (!string.IsNullOrEmpty(ClassFilter))
                {
                    result = result.Union(uiLog.Where(x => x.ClassName.ToLower().Contains(ClassFilter.ToLower())));
                    hadToFilter = true;
                }
                if (!string.IsNullOrEmpty(MemberFilter))
                {
                    result = result.Union(uiLog.Where(x => x.MemberName.ToLower().Contains(MemberFilter.ToLower())));
                    hadToFilter = true;
                }
                if (hadToFilter)
                    return new ObservableCollection<LoggerEventArgs>(result);
                else
                    return uiLog;
            }
            set => SetValue(ref uiLog, value);
        }

        public ICollection<PowerPointPresentationItem> PowerPointPresentations { get; set; }
        public ICollection<IGeneratorOption> GeneratorOptions { get; set; }

        private ICommand buttonSelectTemplateClickCommand;
        public ICommand ButtonSelectTemplateClickCommand
        {
            get
            {
                buttonSelectTemplateClickCommand = buttonSelectTemplateClickCommand ?? new RelayCommand(() =>
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
                buttonSelectFolderClickCommand = buttonSelectFolderClickCommand ?? new RelayCommand(() =>
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
                buttonStartClickCommand = buttonStartClickCommand ?? new RelayCommand(() =>
                {
                    generatorWorker.RunWorkerAsync();
                });
                return buttonStartClickCommand;
            }
        }

        private ICommand buttonResetClickCommand;
        public ICommand ButtonResetClickCommand
        {
            get
            {
                buttonResetClickCommand = buttonResetClickCommand ?? new RelayCommand(() =>
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
                buttonSelectAllPresentationsClickCommand = buttonSelectAllPresentationsClickCommand ?? new RelayCommand(() =>
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
                buttonUnselectAllPresentationsClickCommand = buttonUnselectAllPresentationsClickCommand ?? new RelayCommand(() =>
                {
                    foreach (PowerPointPresentationItem item in PowerPointPresentations)
                        item.IsIncluded = false;
                });
                return buttonUnselectAllPresentationsClickCommand;
            }
        }

        private ICommand filterTextChangedCommand;
        public ICommand FilterTextChangedCommand
        {
            get
            {
                filterTextChangedCommand = filterTextChangedCommand ?? new RelayCommand(() =>
                {
                    OnPropertyChanged(nameof(UILog));
                });
                return filterTextChangedCommand;
            }
        }

        private ICommand exportLogCommand;
        public ICommand ExportLogCommand
        {
            get
            {
                exportLogCommand = exportLogCommand ?? new RelayCommand(() =>
                {
                    //TODO: auslagern auf Klasse?
                    string[] log = uiLog.Select(x => $"{x.Time.ToLongTimeString()};{x.ClassName};{x.MemberName};{x.Message}").ToArray();
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "CSV|*.csv";

                    if(sfd.ShowDialog() == DialogResult.OK)
                    {
                        File.WriteAllLines(sfd.FileName, log);
                        MessageBox.Show("Log wurde erfolgreich exportiert !");
                    }
                });
                return exportLogCommand;
            }
        }

        public void Cleanup()
        {
            ButtonResetClickCommand?.Execute(null);
        }
    }
}
