using Microsoft.Win32;
using ppedv.pocgen.Logic;
using ppedv.pocgen.UI.WPF.Helpers;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using WPFFolderBrowser;

namespace ppedv.pocgen.UI.WPF.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        public MainViewModel()
        {
            IsValidPresentationRootFolderSelected = false;
            UIElementsEnabled = true;
            PowerPointPresentations = new ObservableCollection<PowerPointPresentationItem>();

            tempPath = Path.Combine(Path.GetTempPath(), "pocgen");
            tempImagePath = Directory.CreateDirectory(Path.Combine(tempPath, "genSlides")).FullName;

            // Cleanup
            Directory.Delete(tempImagePath, true);
            Directory.CreateDirectory(tempImagePath);
        }

        private readonly string tempPath;
        private readonly string tempImagePath;

        private string presentationRootFolderPath;
        public string PresentationRootFolderPath
        {
            get => presentationRootFolderPath;
            set
            {
                if (string.IsNullOrWhiteSpace(value) || !Directory.Exists(value)) // Reset
                {
                    SetValue(ref presentationRootFolderPath, string.Empty);
                    IsValidPresentationRootFolderSelected = false;
                    return;
                }
                SetValue(ref presentationRootFolderPath, value);
                // Cleanup
                PowerPointPresentations.Clear();
                Directory.Delete(tempImagePath, true);
                Directory.CreateDirectory(tempImagePath);

                using (PowerPointHelper pph = new PowerPointHelper())
                {
                    int presentationCounter = 0;
                    int presentationStartingIndex = 0;
                    foreach (var file in Directory.GetFiles(presentationRootFolderPath, "*.pptx", SearchOption.AllDirectories).Where(x => x.Contains("~$") == false))
                    {
                        var ppi = new PowerPointPresentationItem(file);
                        ppi.PropertyChanged += (sender, e) =>
                        {
                            if (e.PropertyName == "IsIncluded")
                                IsAtLeastOnePresentationSelected = PowerPointPresentations.Any(x => x.IsIncluded);
                        };

                        ppi.PreviewImagePath = Path.Combine(tempImagePath,presentationCounter++.ToString());
                        // Generate Images (Task ?)
                        var presentation = pph.OpenPresentation(file);
                        Directory.CreateDirectory(ppi.PreviewImagePath);
                        pph.ExportAllSlidesAsImage(presentation, ppi.PreviewImagePath);
                        presentation.Close();

                        int numberOfSlides = Directory.GetFiles(ppi.PreviewImagePath).Length;
                        ppi.PreviewImageRange = (presentationStartingIndex, presentationStartingIndex + numberOfSlides -1);
                        presentationStartingIndex += numberOfSlides;

                        PowerPointPresentations.Add(ppi);
                    }
                }
                if (PowerPointPresentations.Count > 0)
                {
                    int filenumber = 0;
                    foreach(string subdir in Directory.GetDirectories(tempImagePath))
                    {
                        foreach (string file in Directory.GetFiles(subdir))
                        {
                            File.Move(file, Path.Combine(tempImagePath, $"{filenumber++}.png"));
                        }
                    }

                    foreach (string subdir in Directory.GetDirectories(tempImagePath))
                        Directory.Delete(subdir);

                    IsValidPresentationRootFolderSelected = true;
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Valid FolderPath selected");
                }
                else
                {
                    IsValidPresentationRootFolderSelected = false;
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Invalid FolderPath selected");
                }
            }
        }

        private bool isValidFolderSelected;
        public bool IsValidPresentationRootFolderSelected
        {
            get => isValidFolderSelected;
            set => SetValue(ref isValidFolderSelected, value);
        }

        private bool isAtLeastOnePresentationSelected;
        public bool IsAtLeastOnePresentationSelected
        {
            get => isAtLeastOnePresentationSelected;
            set => SetValue(ref isAtLeastOnePresentationSelected, value);
        }

        private bool generatorIsWorking;
        public bool GeneratorIsWorking
        {
            get => generatorIsWorking;
            set => SetValue(ref generatorIsWorking, value);
        }

        private bool uiElementsEnabled;
        public bool UIElementsEnabled
        {
            get => uiElementsEnabled;
            set => SetValue(ref uiElementsEnabled, value);
        }

        private ObservableCollection<PowerPointPresentationItem> powerPointPresentations;
        public ObservableCollection<PowerPointPresentationItem> PowerPointPresentations
        {
            get => powerPointPresentations;
            set => SetValue(ref powerPointPresentations, value);
        }

        private ICommand buttonSelectRootFolderClickCommand;
        public ICommand ButtonSelectRootFolderClickCommand
        {
            get
            {
                buttonSelectRootFolderClickCommand = buttonSelectRootFolderClickCommand ?? new RelayCommand(parameter =>
                {
                    WPFFolderBrowserDialog dlg = new WPFFolderBrowserDialog();
                    if (dlg.ShowDialog() == true)
                    {
                        if (Directory.GetFiles(dlg.FileName, "*.pptx", SearchOption.AllDirectories).Count() > 0)
                            PresentationRootFolderPath = dlg.FileName;
                        else
                        {
                            Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Directory does not contain .pptx files.");
                            MessageBox.Show("Das Verzeichnis  beinhaltet keine .pptx - Dateien");
                        }
                    }
                });
                return buttonSelectRootFolderClickCommand;
            }
        }

        private ICommand buttonStartClickCommand;
        public ICommand ButtonStartClickCommand
        {
            get
            {
                buttonStartClickCommand = buttonStartClickCommand ?? new RelayCommand(parameter =>
                {
                    UIElementsEnabled = false;
                    GeneratorIsWorking = true;
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Generator-Start");

                    string tempPresentation = Path.Combine(tempPath, $"{Guid.NewGuid()}.pptx");
                    string tempDocument = Path.Combine(tempPath, $"{Guid.NewGuid()}.docx");

                    using (PowerPointHelper pph = new PowerPointHelper())
                    {
                        var mergedPresentation = pph.CreateNewPresentation(tempPresentation);
                        pph.MergePresentationContentIntoNewPresentation(PowerPointPresentations.Where(x => x.IsIncluded).Select(x => x.FullPath), mergedPresentation);
                        pph.ExportAllSlidesAsImage(mergedPresentation, tempImagePath);
                        pph.SavePresentationAs(mergedPresentation, tempPresentation);
                    }

                    WordGenerator generator = new WordGenerator();
                    generator.GeneratePOC_Document(tempPresentation, tempImagePath, tempDocument);
                    Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Generator-Finish");

                    SaveFileDialog dlg = new SaveFileDialog();
                    dlg.Title = "POC Speichern unter";
                    dlg.Filter = "Word Dokument | *.docx";

                    if (dlg.ShowDialog() == true)
                    {
                        File.Move(tempDocument, dlg.FileName);
                    }

                    UIElementsEnabled = true;
                    GeneratorIsWorking = false;
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
                    IsValidPresentationRootFolderSelected = false;
                    PowerPointPresentations.Clear();
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
    }
}
