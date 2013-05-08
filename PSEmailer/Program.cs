using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using SendGridMail;
using SendGridMail.Transport;
using UserInput;

namespace PSEmailer
{
    internal class Program
    {
        private static string _printerOutputFullPath;

        private static Task _gsTask;

        private static void Main()
        {
            var uiThread = new Thread(UiThread);
            uiThread.SetApartmentState(ApartmentState.STA);
            uiThread.Start();

            string tempFile = Path.GetTempFileName();
            
            using (Stream stdin = Console.OpenStandardInput())
            {
                using (FileStream stdout = File.OpenWrite(tempFile))
                {
                    stdin.CopyTo(stdout);
                }
            }
              
            _gsTask = new Task(() => PDFConversion(tempFile));
            _gsTask.Start();
        }

        public static void UiThread(object arg)
        {
            var app = new Application();
            var window = new MainWindow();
            window.onClosed += window_onClosed;
            app.Run(window);
        }

        private static void PDFConversion(string tempFile)
        {
            const string adAssistPrinterFolder = "AdAssistPrinter";
            const string printerOutput = "PrintFromADAssist.pdf";

            string localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            string adAssistPrinterFolderFullPath = Path.Combine(localApplicationData, adAssistPrinterFolder);
            if (!Directory.Exists(adAssistPrinterFolderFullPath))
            {
                Directory.CreateDirectory(adAssistPrinterFolderFullPath);
            }

            string printerOutputFullPath = Path.Combine(adAssistPrinterFolderFullPath, printerOutput);
            if (File.Exists(printerOutputFullPath))
            {
                File.Delete(printerOutputFullPath);
            }

            _printerOutputFullPath = printerOutputFullPath;

            var gs = new GhostScript();
            gs.AddParam("-q");
            gs.AddParam("-dNOPAUSE");
            gs.AddParam("-dBATCH");
            gs.AddParam("-dQUIET");
            gs.AddParam("-sDEVICE=pdfwrite");
            gs.AddParam("-dSAFER");
            gs.AddParam("-sPAPERSIZE=letter");
            gs.AddParam("-sOutputFile=" + _printerOutputFullPath);
            gs.AddParam(tempFile);
            gs.Execute();
        }

        private static void window_onClosed(object sender, ClosedArgs args)
        {
            _gsTask.Wait();

            SendGrid mail = SendGrid.GetInstance();

            mail.AddTo(args.EmailAddress);
            mail.From = new MailAddress("no-reply@8to18.com", "8to18, Inc.");
            mail.Subject = "Your Printout From ADAssist";
            mail.Text = "Your printout from ADAssist is attached.";
            mail.AddAttachment(_printerOutputFullPath);

            var credentials = new NetworkCredential("8to18media", "YpiRBTzEfOyszKOGcRTZ");
            Web transportWeb = Web.GetInstance(credentials);

            transportWeb.Deliver(mail);

            File.Delete(_printerOutputFullPath);
            _printerOutputFullPath = string.Empty;
        }
    }
}