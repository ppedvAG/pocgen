using System;
using System.Text;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Input;
using System.Threading.Tasks;
using ppedv.pocgen.UI.WPF.Helpers;
using WPFFolderBrowser;
using ppedv.pocgen.Logic;
using Microsoft.Win32;
using System.Windows;

namespace ppedv.pocgen.UI.WPF.ViewModels
{
    public class MainViewModel : BaseViewModel
    {
        public MainViewModel()
        {
            IsValidPresentationRootFolderSelected = false;
            UIElementsEnabled = true;
            PowerPointPresentations = new ObservableCollection<PowerPointPresentationItem>();

            tempPath =  Path.Combine(Path.GetTempPath(),"pocgen");
            tempImagePath = Directory.CreateDirectory(Path.Combine(tempPath, "genSlides")).FullName;
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
                PowerPointPresentations.Clear();

                foreach (var file in Directory.GetFiles(presentationRootFolderPath, "*.pptx", SearchOption.AllDirectories))
                    PowerPointPresentations.Add(new PowerPointPresentationItem(file));

                if (PowerPointPresentations.Count > 0)
                {
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
                    Task.Run(() =>
                    {
                        UIElementsEnabled = false;
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
