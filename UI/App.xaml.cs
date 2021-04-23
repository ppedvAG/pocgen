using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace pocgen
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        public string SEMID { get; private set; }
        public string KursName { get; private set; }
        public string UploadURL { get; private set; }

        private void Application_Startup(object sender, StartupEventArgs e)
        {
            try
            {
                if (e.Args.Count() > 0)
                {
                    var chunks = e.Args.First().Split('|');
                    if (chunks.Count() == 3)
                    {
                        SEMID = chunks[0];
                        KursName = chunks[1];
                        UploadURL = chunks[2];
                    }
                    else
                    {
                        ShowCommandLineInfo();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Application_Startup exception: {ex.Message}");
            }
        }

        public static void ShowCommandLineInfo()
        {
            var info = @"Für den Upload per PocGen muss ein einzelner Startparameter angegeben werden.
Die Daten müssen mit einem senkrechten Strich (|) getrennt und in der korrekten Reihenfolge angegeben werden:

    SEMID|KURSNAME|Upload URL

z.B. 12345|Upload für Anfänger|https://ppedv.de/AnyPage.ashx

Die echte Upload-URL wird nicht in Programm hinterlegt, weil der Quellcode öffentlich zugänglich ist.";

            MessageBox.Show(info,"Startparameter");
        }
    }

}
