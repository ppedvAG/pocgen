using AdonisUI;
using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
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

namespace pocgen
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private bool _isDark;

        private void ChangeTheme(object sender, RoutedEventArgs e)
        {
            ResourceLocator.SetColorScheme(Application.Current.Resources, _isDark ? ResourceLocator.LightColorScheme : ResourceLocator.DarkColorScheme);
            _isDark = !_isDark;
        }

        private void ShowInfoText(object sender, RoutedEventArgs e)
        {
            MessageBox.Show($"Anleitung:{Environment.NewLine}" +
                $"Im ersten Schritt müssen Sie mithilfe der Schaltfläche 'Ordner auswählen' einen Ordner wählen, der zumindest eine PowerPoint-Präsentation beinhaltet. {Environment.NewLine}" +
                $"Danach können Sie in der unteren Liste die Präsentationen auswählen, die vom Programm bearbeitet werden. Eine Vorschau auf der rechten Seite zeigt an, welche Folien verarbeitet werden.{Environment.NewLine}" +
                $"Zuletzt müssen Sie eine der folgenden 4 Aktionen für die Verarbeitung wählen:{Environment.NewLine}" +
                $"1) Aus allen Präsentationen ein Word-Dokument generieren und als Word-Dokument speichern{Environment.NewLine}" +
                $"Mit dieser Aktion wird, wie in den vorherigen Versionen von pocgen üblich, ein Word-Dokument mit dem Folieninhalt und den Notizen der jeweiligen Folie generiert.{Environment.NewLine}" +
                $"2) Aus allen Präsentationen ein Word-Dokument generieren und als PDF speichern{Environment.NewLine}" +
                $"Mit dieser Aktion wird das Selbe wie in Aktion 1) gemacht, nur dass das Ergebnis als PDF und nicht als Word-Datei gespeichert wird.{Environment.NewLine}" +
                $"3) Alle Präsentationen zu einer einzelnen PowerPoint-Präsentation zusammenfassen{Environment.NewLine}" +
                $"Mit dieser Aktion wird eine neue PowerPoint-Präsentation erstellt, deren Inhalt aus allen ausgewählten Präsentationen besteht. Die einzelnen Dateien bleiben hierbei erhalten.{Environment.NewLine}" +
                $"4) Aus allen Präsentationen eine PDF-Datei generieren{Environment.NewLine}" +
                $"Mit dieser Aktion wird das Selbe wie in Aktion 3) gemacht, nur dass das Ergebnis als PDF und nicht als PowerPoint-Präsentation gespeichert wird.");
        }

        private void ShowAboutText(object sender, RoutedEventArgs e)
        {
            string version = "DEBUG";
            if (ApplicationDeployment.IsNetworkDeployed)
                version = $"{ApplicationDeployment.CurrentDeployment.CurrentVersion.Major}.{ApplicationDeployment.CurrentDeployment.CurrentVersion.Minor}.{ApplicationDeployment.CurrentDeployment.CurrentVersion.Build}.{ApplicationDeployment.CurrentDeployment.CurrentVersion.Revision}";

            MessageBox.Show($"ppedv official course generator{Environment.NewLine}Version: {version}");
        }
    }
}
