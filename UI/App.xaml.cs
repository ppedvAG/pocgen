using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
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
                    var argAsUri = Uri.UnescapeDataString(e.Args.First());
                    var chunks = argAsUri.ToString().Split('|');
                    if (chunks.Count() == 3)
                    {
                        SEMID = chunks[0].Split('/')[2];
                        KursName = chunks[1];
                        UploadURL = chunks[2];
                    }
                    else
                    {
                        ShowCommandLineInfo(e.Args[0]);
                    }
                }
            }
            catch (Exception ex)
            {
                ShowCommandLineInfo(e.Args[0], $"Application_Startup exception: {ex.Message}");
            }
        }

        public static void ShowCommandLineInfo(string startParameter = "", string errorMsg = "")
        {
            var info = @"Für den Upload per PocGen muss ein einzelner Startparameter angegeben werden.
Die Daten müssen mit einem senkrechten Strich (|) getrennt und in der korrekten Reihenfolge angegeben werden:

    SEMID|KURSNAME|Upload URL

z.B. 12345|Upload für Anfänger|https://ppedv.de/AnyPage.ashx

Die echte Upload-URL wird nicht in Programm hinterlegt, weil der Quellcode öffentlich zugänglich ist.";


            var spInfo = $"Startparameter: {startParameter}";

            var sb = new StringBuilder();
            if (!string.IsNullOrEmpty(startParameter))
            {
                sb.AppendLine(spInfo);
                sb.AppendLine();
            }

            if (!string.IsNullOrWhiteSpace(errorMsg))
            {
                sb.AppendLine(errorMsg);
                sb.AppendLine();
            }

            sb.Append(info);

            MessageBox.Show(sb.ToString(), "Startparameter");
        }
    }

}
