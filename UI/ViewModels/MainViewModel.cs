using Microsoft.Win32;
using ppedv.pocgen.Logic;
using ppedv.pocgen.UI.WPF.Helpers;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
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

            // Cleanup
            tempPath = Path.Combine(Path.GetTempPath(), "pocgen");
            if (Directory.Exists(tempPath))
                Directory.Delete(tempPath, true);
            Directory.CreateDirectory(tempPath);

            tempImagePathForPOC = Directory.CreateDirectory(Path.Combine(tempPath, "genSlidesPOC")).FullName;
            IsGeneratingPreview = false;

            SelectPresentationSourceForUploadCommand = new RelayCommand(o => SelectPresentationForUpload());
            SelectSampleSourceForUploadCommand = new RelayCommand(o => SelectSampleSourceForUpload());
            UploadCommand = new RelayCommand(o => Upload());
        }



        private readonly string tempPath;
        private string tempImagePath;
        private readonly string tempImagePathForPOC;

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
                CurrentSlide = 0;
                PowerPointPresentations.Clear();

                tempImagePath = Path.Combine(tempPath, Guid.NewGuid().ToString());

                Task.Run(() =>
                {
                    IsGeneratingPreview = true;
                    using (PowerPointHelper pph = new PowerPointHelper())
                    {
                        int presentationCounter = 0;
                        int presentationStartingIndex = 0;
                        foreach (var file in Directory.GetFiles(presentationRootFolderPath, "*.pptx", SearchOption.AllDirectories).Where(x => x.Contains("~$") == false).OrderBy(f => f))
                        {
                            var ppi = new PowerPointPresentationItem(file);
                            ppi.PropertyChanged += (sender, e) =>
                            {
                                if (e.PropertyName == "IsIncluded")
                                    IsAtLeastOnePresentationSelected = PowerPointPresentations.Any(x => x.IsIncluded);
                            };

                            ppi.PreviewImagePath = Path.Combine(tempImagePath, presentationCounter++.ToString().PadLeft(5, '0'));
                            // Generate Images (Task ?)
                            var presentation = pph.OpenPresentation(file);
                            Directory.CreateDirectory(ppi.PreviewImagePath);
                            pph.ExportAllSlidesAsImage(presentation, ppi.PreviewImagePath);
                            presentation.Close();

                            int numberOfSlides = Directory.GetFiles(ppi.PreviewImagePath).Length;
                            ppi.PreviewImageRange = (presentationStartingIndex, presentationStartingIndex + numberOfSlides - 1);
                            presentationStartingIndex += numberOfSlides;

                            Application.Current.Dispatcher.Invoke(() => PowerPointPresentations.Add(ppi));
                        }
                    }
                    if (PowerPointPresentations.Count > 0)
                    {
                        int filenumber = 0;
                        foreach (string subdir in Directory.GetDirectories(tempImagePath))
                        {
                            foreach (string file in Directory.GetFiles(subdir))
                            {
                                File.Move(file, Path.Combine(tempImagePath, $"{filenumber++}.png"));
                            }
                        }

                        foreach (string subdir in Directory.GetDirectories(tempImagePath))
                            Directory.Delete(subdir);

                        ResetPreviewCommand.Execute(null);
                        IsValidPresentationRootFolderSelected = true;
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Valid FolderPath selected");
                    }
                    else
                    {
                        IsValidPresentationRootFolderSelected = false;
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Invalid FolderPath selected");
                    }
                    IsGeneratingPreview = false;
                });
            }
        }

        private bool isGeneratingPreview;
        public bool IsGeneratingPreview
        {
            get => isGeneratingPreview;
            set => SetValue(ref isGeneratingPreview, value);
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
                            //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Directory does not contain .pptx files.");
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
                        GeneratorIsWorking = true;
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Generator-Start");

                        string tempPresentation = Path.Combine(tempPath, $"{Guid.NewGuid()}.pptx");
                        string tempDocument = Path.Combine(tempPath, $"{Guid.NewGuid()}.docx");

                        using (PowerPointHelper pph = new PowerPointHelper())
                        {
                            var mergedPresentation = pph.CreateNewPresentation();
                            pph.MergePresentationContentIntoNewPresentation(PowerPointPresentations.Where(x => x.IsIncluded).Select(x => x.FullPath), mergedPresentation);
                            pph.ExportAllSlidesAsImage(mergedPresentation, tempImagePathForPOC);
                            pph.SavePresentationAs(mergedPresentation, tempPresentation);
                            mergedPresentation.Close();
                        }

                        WordGenerator generator = new WordGenerator();
                        generator.GeneratePOC_Document(tempPresentation, tempImagePathForPOC, tempDocument);
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Generator-Finish");

                        SaveFileDialog dlg = new SaveFileDialog();
                        dlg.Title = "POC Speichern unter";
                        dlg.Filter = "Word Dokument | *.docx";

                        if (dlg.ShowDialog() == true)
                        {
                            if (File.Exists(dlg.FileName))
                                File.Delete(dlg.FileName);

                            File.Move(tempDocument, dlg.FileName);
                        }

                        UIElementsEnabled = true;
                        GeneratorIsWorking = false;
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

        private ICommand mergePresentationsCommand;
        public ICommand MergePresentationsCommand
        {
            get
            {
                mergePresentationsCommand = mergePresentationsCommand ?? new RelayCommand(parameter =>
                {
                    Task.Run(() =>
                    {
                        UIElementsEnabled = false;
                        GeneratorIsWorking = true;
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PPTX-Generator-Start");

                        string tempPresentation = Path.Combine(tempPath, $"{Guid.NewGuid()}.pptx");
                        string tempDocument = Path.Combine(tempPath, $"{Guid.NewGuid()}.docx");

                        using (PowerPointHelper pph = new PowerPointHelper())
                        {
                            var mergedPresentation = pph.CreateNewPresentation();
                            pph.MergePresentationContentIntoNewPresentation(PowerPointPresentations.Where(x => x.IsIncluded).Select(x => x.FullPath), mergedPresentation);
                            pph.ExportAllSlidesAsImage(mergedPresentation, tempImagePathForPOC);
                            pph.SavePresentationAs(mergedPresentation, tempPresentation);
                            mergedPresentation.Close();
                        }

                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PPTX-Generator-Finish");

                        SaveFileDialog dlg = new SaveFileDialog();
                        dlg.Title = "PPTX Speichern unter";
                        dlg.Filter = "PPTX| *.pptx";

                        if (dlg.ShowDialog() == true)
                        {
                            if (File.Exists(dlg.FileName))
                                File.Delete(dlg.FileName);

                            File.Move(tempPresentation, dlg.FileName);
                        }

                        UIElementsEnabled = true;
                        GeneratorIsWorking = false;
                    });
                });
                return mergePresentationsCommand;
            }
        }

        private ICommand generatePresentationPDFCommand;
        public ICommand GeneratePresentationPDFCommand
        {
            get
            {
                generatePresentationPDFCommand = generatePresentationPDFCommand ?? new RelayCommand(parameter =>
                {
                    Task.Run(() =>
                    {
                        UIElementsEnabled = false;
                        GeneratorIsWorking = true;
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PPTX_to_PDF-Generator-Start");

                        string tempPresentation = Path.Combine(tempPath, $"{Guid.NewGuid()}.pptx");
                        string tempDocument = Path.Combine(tempPath, $"{Guid.NewGuid()}.docx");

                        using (PowerPointHelper pph = new PowerPointHelper())
                        {
                            var mergedPresentation = pph.CreateNewPresentation();
                            pph.MergePresentationContentIntoNewPresentation(PowerPointPresentations.Where(x => x.IsIncluded).Select(x => x.FullPath), mergedPresentation);
                            pph.ExportAllSlidesAsImage(mergedPresentation, tempImagePathForPOC);
                            pph.SavePresentationAs(mergedPresentation, tempPresentation);

                            //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PPTX_to_PDF-Generator-Finish");

                            SaveFileDialog dlg = new SaveFileDialog();
                            dlg.Title = "PDF Speichern unter";
                            dlg.Filter = "PDF| *.pdf";

                            if (dlg.ShowDialog() == true)
                            {
                                if (File.Exists(dlg.FileName))
                                    File.Delete(dlg.FileName);

                                // ToDo: In den Helper auslagern
                                mergedPresentation.ExportAsFixedFormat(dlg.FileName, Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                            }
                            mergedPresentation.Close();
                            PresentationSourceURI = dlg.FileName;
                        }
                        UIElementsEnabled = true;
                        GeneratorIsWorking = false;
                    });
                });
                return generatePresentationPDFCommand;
            }
        }

        private ICommand generatePOC_PDFCommand;
        public ICommand GeneratePOC_PDFCommand
        {
            get
            {
                generatePOC_PDFCommand = generatePOC_PDFCommand ?? new RelayCommand(parameter =>
                {
                    Task.Run(() =>
                    {
                        UIElementsEnabled = false;
                        GeneratorIsWorking = true;
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PDF-Generator-Start");

                        string tempPresentation = Path.Combine(tempPath, $"{Guid.NewGuid()}.pptx");
                        string tempDocument = Path.Combine(tempPath, $"{Guid.NewGuid()}.docx");

                        using (PowerPointHelper pph = new PowerPointHelper())
                        {
                            var mergedPresentation = pph.CreateNewPresentation();
                            pph.MergePresentationContentIntoNewPresentation(PowerPointPresentations.Where(x => x.IsIncluded).Select(x => x.FullPath), mergedPresentation);
                            pph.ExportAllSlidesAsImage(mergedPresentation, tempImagePathForPOC);
                            pph.SavePresentationAs(mergedPresentation, tempPresentation);
                            mergedPresentation.Close();
                        }

                        WordGenerator generator = new WordGenerator();
                        generator.GeneratePOC_Document(tempPresentation, tempImagePathForPOC, tempDocument);
                        //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] PDF-Generator-Finish");

                        SaveFileDialog dlg = new SaveFileDialog();
                        dlg.Title = "PDF Speichern unter";
                        dlg.Filter = "PDF| *.pdf";

                        if (dlg.ShowDialog() == true)
                        {
                            if (File.Exists(dlg.FileName))
                                File.Delete(dlg.FileName);

                            // ToDo: Auslagern in die Logik-Layer
                            // Word öffnen, PDF erzeugen und an dem Ort speichern
                            Microsoft.Office.Interop.Word.Application appWD = new Microsoft.Office.Interop.Word.Application();
                            var wordDocument = appWD.Documents.Open(tempDocument);
                            wordDocument.ExportAsFixedFormat(dlg.FileName, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                            wordDocument.Close(false);
                            wordDocument = null;
                            appWD.Quit();
                            appWD = null;
                        }

                        UIElementsEnabled = true;
                        GeneratorIsWorking = false;
                    });
                });
                return generatePOC_PDFCommand;
            }
        }
        #region Preview
        private string[] includedImages;

        private ImageSource previewSource;
        public ImageSource PreviewSource
        {
            get => previewSource;
            set => SetValue(ref previewSource, value);
        }
        private int currentSlide;
        public int CurrentSlide
        {
            get => currentSlide;
            set
            {
                SetValue(ref currentSlide, value);
                OnPropertyChanged(nameof(IsMaximumReached));
                if (includedImages == null || includedImages.Length == 0 || CurrentSlide == 0)
                {
                    PreviewSource = null;
                    return;
                }
                try
                {
                    PreviewSource = new BitmapImage(new Uri(includedImages[CurrentSlide - 1]));
                }
                catch (FileNotFoundException)
                {
                    //Trace.WriteLine($"[{GetType().Name}|{MethodBase.GetCurrentMethod().Name}] Exception: Slide {CurrentSlide - 1}.png not found !");
                    PreviewSource = null;
                }
            }
        }
        private int maximumSlides;
        public int MaximumSlides
        {
            get => maximumSlides;
            set => SetValue(ref maximumSlides, value);
        }

        public bool IsMaximumReached
        {
            get => CurrentSlide == MaximumSlides;
        }

        private ICommand resetPreviewCommand;
        public ICommand ResetPreviewCommand
        {
            get
            {
                resetPreviewCommand = resetPreviewCommand ?? new RelayCommand(parameter =>
                {
                    if (PowerPointPresentations == null || PowerPointPresentations.Count() == 0)
                    {
                        CurrentSlide = 1;
                        MaximumSlides = 1;
                        return;
                    }

                    MaximumSlides = PowerPointPresentations.Where(x => x.IsIncluded)
                                                           .Sum(x => x.PreviewImageRange.Item2 + 1 - x.PreviewImageRange.Item1);

                    includedImages = PowerPointPresentations.Where(x => x.IsIncluded)
                                                            .SelectMany(x => Enumerable.Range(x.PreviewImageRange.Item1, (x.PreviewImageRange.Item2 + 1) - x.PreviewImageRange.Item1))
                                                            .Select(x => System.IO.Path.Combine(tempImagePath, $"{x}.png"))
                                                            .ToArray();
                    CurrentSlide = MaximumSlides == 0 ? 0 : 1;
                });
                return resetPreviewCommand;
            }
        }

        private ICommand previewForwardCommand;
        public ICommand PreviewForwardCommand
        {
            get
            {
                previewForwardCommand = previewForwardCommand ?? new RelayCommand(parameter =>
                {
                    if (CurrentSlide < MaximumSlides)
                        CurrentSlide++;
                });
                return previewForwardCommand;
            }
        }
        private ICommand previewBackwardCommand;
        private bool isSampleFileUploadSelected;
        private string sampleSourceURI;
        private string presentationSourceURI;

        public ICommand PreviewBackwardCommand
        {
            get
            {
                previewBackwardCommand = previewBackwardCommand ?? new RelayCommand(parameter =>
                {
                    if (CurrentSlide != 1)
                        CurrentSlide--;
                });
                return previewBackwardCommand;
            }
        }

        #endregion

        public bool IsSampleFileUploadSelected
        {
            get => isSampleFileUploadSelected;
            set
            {
                isSampleFileUploadSelected = value;
                OnPropertyChanged();
                OnPropertyChanged(nameof(IsSampleLinkSelected));
            }
        }
        public bool IsSampleLinkSelected { get => !IsSampleFileUploadSelected; }

        public string SampleSourceURI
        {
            get => sampleSourceURI;
            set
            {
                sampleSourceURI = value;
                OnPropertyChanged();
            }
        }

        public string PresentationSourceURI
        {
            get => presentationSourceURI;
            set
            {
                presentationSourceURI = value;
                OnPropertyChanged();
            }
        }

        public ICommand SelectPresentationSourceForUploadCommand { get; private set; }
        public ICommand SelectSampleSourceForUploadCommand { get; private set; }
        public ICommand UploadCommand { get; private set; }

        private void SelectPresentationForUpload()
        {
            var dlg = new OpenFileDialog() { FileName = PresentationSourceURI, Filter = "ZIP-Datei oder Präsentation|*.zip;*.pptx", Multiselect = false };

            if (dlg.ShowDialog().Value)
            {
                PresentationSourceURI = dlg.FileName;
            }

        }

        private void SelectSampleSourceForUpload()
        {
            var dlg = new OpenFileDialog() { Filter = "ZIP-Datei|*.zip|Alle Dateiarten|*.*", Multiselect = false };

            if (dlg.ShowDialog().Value)
            {
                SampleSourceURI = dlg.FileName;
            }
        }

        private async void Upload()
        {

            var url = $"https://download.ppedv.de/FileUploadHandler.ashx";
            
            
            //todo change
            string SemAppId = "208493";
            var sampleURL = "http://www.github.com/ppedvag";
            var presFileName = "Presentation_208493_Tests.zip";
            var presFilePath = @"C:\Users\rulan\Desktop\Roßberger Upload Test\Presentation_208493_Tests.zip";
            var sampFileName = "Samples_208493_url.txt";
            var sampFilePath = @"C:\Users\rulan\Desktop\Roßberger Upload Test\Samples_208493_url.txt";

            var http = new HttpClient();

            var httpContent = new MultipartFormDataContent();
            httpContent.Add(new StringContent(sampleURL), "SampleUrl");
            httpContent.Add(new StringContent(SemAppId), "SemAppId");
            httpContent.Add(new ByteArrayContent(File.ReadAllBytes(presFilePath)), $"{presFileName}_P", presFileName);
            httpContent.Add(new ByteArrayContent(File.ReadAllBytes(sampFilePath)), $"{sampFileName}_S", sampFileName);

            var response = await http.PostAsync(url, httpContent);

            if (response.IsSuccessStatusCode)
                MessageBox.Show("Vielen Dank für den Upload, Sie dürfen sich nun einen Keks genehmigen und Ihre nächste Stufe der Existenz genießen.");
            else
                MessageBox.Show($"ERROR: {(int)response.StatusCode} {response.ReasonPhrase}");
        }

    }
}
