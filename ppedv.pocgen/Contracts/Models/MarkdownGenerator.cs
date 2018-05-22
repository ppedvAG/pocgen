using Markdig;
using pocgen.Contracts.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace pocgen.Contracts.Models
{
    public class MarkdownGenerator : IMarkdownGenerator
    {
        public MarkdownGenerator()
        {
            webbrowser = new WebBrowser();
            webbrowser.CreateControl();
        }

        private WebBrowser webbrowser;

        public void GenerateMarkdownFromInputAndCopyIntoClipboard(string input)
        {
            // BÖSER HACK  ! -> Mit CopyPaste kann das formatierte Markdown aus einem Webbrowser in Word hineinkopiert werden
            webbrowser.Invoke(new Action(() =>
            {
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Start of generating Markdown"));
               webbrowser.Navigate("about:blank");
                while (webbrowser.ReadyState != WebBrowserReadyState.Complete) // Warten bis die Navigation beendet wurde, ansonsten funktioniert das "Clear" der Webseite nicht und wir bekommen doppelten Text raus
                {
                    Application.DoEvents(); // Ohne dem wird der Webbrowser nicht die neue seite "laden"
                    Thread.Sleep(100);
                }

                webbrowser.Document.Write(Markdown.ToHtml(input));
                webbrowser.Document.ExecCommand("SelectAll", false, null);
                webbrowser.Document.ExecCommand("Copy", false, null);
                MessagingCenter.Send(this, "Log", new LoggerEventArgs(GetType().Name, MethodBase.GetCurrentMethod().Name, $"Markdown copied to clipboard"));
            })); ;
        }
    }
}
