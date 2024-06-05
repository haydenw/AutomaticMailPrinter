using MailKit.Net.Imap;
using MailKit;
using System;
using MailKit.Search;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using DinkToPdf;
using Orientation = DinkToPdf.Orientation;
using PdfiumViewer;
using System.Drawing.Printing;
using PaperKind = DinkToPdf.PaperKind;
using System.Drawing;
using System.Text.RegularExpressions;

namespace AutomaticMailPrinter
{
    internal class Program
    {
        private static System.Threading.Timer timer;
        private static readonly Mutex AppMutex = new Mutex(false, "c75adf4e-765c-4529-bf7a-90dd76cd386a");

        private static string ImapServer, MailAddress, Password, PrinterName;
        public static string WebHookUrl { get; private set; }
        private static string[] Filter = new string[0];
        private static int ImapPort;

        private static ImapClient client = new ImapClient();
        private static IMailFolder inbox;

        private static object sync = new object();

        private static Database database = new Database();

        static void Main(string[] args)
        {
            if (!AppMutex.WaitOne(TimeSpan.FromSeconds(1), false))
            {
                MessageBox.Show(Properties.Resources.strInstanceAlreadyRunning, Properties.Resources.strError, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;
            int intervalInSecods = 60;
            try
            {
                string configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "config.json");
                var configDocument = JsonSerializer.Deserialize<JsonDocument>(File.ReadAllText(configPath));

                ImapServer = configDocument.RootElement.GetProperty("imap_server").GetString();
                ImapPort = configDocument.RootElement.GetProperty("imap_port").GetInt32();
                MailAddress = configDocument.RootElement.GetProperty("mail").GetString();
                Password = configDocument.RootElement.GetProperty("password").GetString();
                PrinterName = configDocument.RootElement.GetProperty("printer_name").GetString();

                try
                {
                    // Can be empty or even may not existing ...
                    WebHookUrl = configDocument.RootElement.GetProperty("webhook_url").GetString();
                }
                catch { }

                intervalInSecods = configDocument.RootElement.GetProperty("timer_interval_in_seconds").GetInt32();

                var filterProperty = configDocument.RootElement.GetProperty("filter");
                int counter = 0;
                Filter = new string[filterProperty.GetArrayLength()];
                foreach (var word in filterProperty.EnumerateArray())
                    Filter[counter++] = word.GetString().ToLower();
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToReadConfigFile, ex);
            }

            try
            {
                Logger.LogInfo(string.Format(Properties.Resources.strConnectToMailServer, $"\"{ImapServer}:{ImapPort}\""));
                client = new ImapClient();
                client.Connect(ImapServer, ImapPort, true);
                client.Authenticate(MailAddress, Password);
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToRecieveMails, ex);
            }
            
            timer = new System.Threading.Timer(Timer_Tick, null, 0, intervalInSecods * 1000);
            GC.KeepAlive(timer);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            using (Form form = new Form())
            {
                NotifyIcon notifyIcon = new NotifyIcon
                {
                    Text = "Order Printer",
                    Icon = Properties.Resources.icon,
                    Visible = true
                };
                ContextMenuStrip contextMenu = new ContextMenuStrip();
                ToolStripMenuItem showOrdersItem = new ToolStripMenuItem("Show Orders");
                showOrdersItem.Click += ShowOrders_Click;
                ToolStripMenuItem exitItem = new ToolStripMenuItem("Exit");
                exitItem.Click += Exit_Click;
                contextMenu.Items.Add(showOrdersItem);
                contextMenu.Items.Add(exitItem);
                notifyIcon.ContextMenuStrip = contextMenu;
                notifyIcon.Visible = true;
                Application.Run();
                notifyIcon.Visible = false;
            }

            AppMutex.ReleaseMutex();
        }

        private static void ShowOrders_Click(object sender, EventArgs e)
        {
            Form formOrders = new FormOrders();
            formOrders.ShowDialog();
        }
        private static void Exit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            if (e.ExceptionObject is Exception ex)
                Logger.LogError("Unhandled Exception recieved!", ex);
            else if (e.ExceptionObject != null)
                Logger.LogError($"Unhandled Exception recieved: {e.ExceptionObject}");
            else
                Logger.LogError("Unhandled Exception but exception object is empty :(");
        }

        private static void Timer_Tick(object state)
        {
            try
            {
                lock (sync)
                {
                    Logger.LogInfo(Properties.Resources.strLookingForUnreadMails);
                    bool found = false;

                    if (!client.IsAuthenticated || !client.IsConnected || inbox == null)
                    {
                        Logger.LogWarning(Properties.Resources.strMailClientIsNotConnectedAnymore);

                        try
                        {
                            client = new ImapClient();
                            client.Connect(ImapServer, ImapPort, true);
                            client.Authenticate(MailAddress, Password);

                            // The Inbox folder is always available on all IMAP servers...
                            inbox = client.Inbox;

                            Logger.LogInfo(Properties.Resources.strConnectionEstablishedSuccess, sendWebHook: true);
                        }
                        catch (Exception ex)
                        {
                            Logger.LogError(Properties.Resources.strFailedToConnect, ex);
                            return;
                        }

                    }

                    inbox.Open(FolderAccess.ReadWrite);

                    foreach (var uid in inbox.Search(SearchQuery.NotSeen))
                    {
                        var message = inbox.GetMessage(uid);
                        string subject = message.Subject.ToLower();
                        if (Filter.Any(f => subject.Contains(f)))
                        {
                            // Extract order number
                            string pattern = @"order #(\d+)";
                            Match match = Regex.Match(subject, pattern);
                            if (!match.Success)
                            {
                                Logger.LogError(string.Format("Failed to extract order number from subject: {0}", subject));
                                continue;
                            }

                            int orderNumber = int.Parse(match.Groups[1].Value);

                            Logger.LogInfo(string.Format("Found order #{0}", orderNumber));

                            try
                            {
                                database.AddOrder(orderNumber, message.HtmlBody, message.Subject);
                            }
                            catch (Exception ex)
                            {
                                Logger.LogError("Failed to add order!", ex);
                            }

                            // Print text
                            Console.ForegroundColor = ConsoleColor.Green;
                            Logger.LogInfo($"{string.Format(Properties.Resources.strFoundUnreadMail, Filter.Where(f => subject.Contains(f)).FirstOrDefault())} {message.Subject}");

                            // Print mail
                            Logger.LogInfo(string.Format(Properties.Resources.strPrintMessage, message.Subject, PrinterName));
                            PrintHtmlPage(orderNumber, message.HtmlBody);

                            Logger.LogInfo("Printed");

                            database.OrderPrinted(orderNumber);

                            Logger.LogInfo("Set order printed");

                            // `Read mail https://stackoverflow.com/a/24204804/6237448
                            Logger.LogInfo(Properties.Resources.strMarkMailAsDeleted);                     
                            //inbox.SetFlags(uid, MessageFlags.Deleted, true);
                            inbox.SetFlags(uid, MessageFlags.Seen, true);

                            Logger.LogInfo("Read email");

                            found = true;

                            PlaySound();
                        }
                    }

                    if (!found)
                        Logger.LogInfo(Properties.Resources.strNoUnreadMailFound);
                    else
                    {
                        Logger.LogInfo(Properties.Resources.strMarkMailAsRead);
                        inbox.Expunge();
                    }

                    // Do not disconnect here!
                    // client.Disconnect(true);
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToRecieveMails, ex);
                Logger.LogError(ex.StackTrace);
            }
        }        

        public static void PlaySound()
        {
            try
            {
                System.Media.SoundPlayer player = new System.Media.SoundPlayer(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Sounds", "sound.wav"));
                player.Play();
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToPlaySound, ex);
            }
        }

        public static void PrintHtmlPage(int orderNumber, string htmlContent)
        {
            try
            {
                string pdfPath = ConvertHtmlToPdf(orderNumber, htmlContent);
                Logger.LogInfo("Converted to PDF");
                // TODO improve unicode support
                htmlContent.Replace("×", "x");
                PrintPdf(pdfPath, PrinterName);
                Logger.LogInfo("Printed pdf");
                File.Delete(pdfPath);
                Logger.LogInfo("Deleted pdf");
            }
            catch (Exception ex)
            {
                Logger.LogError(Properties.Resources.strFailedToPrintMail, ex);
            }
        }

        private static string ConvertHtmlToPdf(int orderNumber, string htmlContent)
        {
            var converter = new SynchronizedConverter(new PdfTools());

            var doc = new HtmlToPdfDocument()
            {
                GlobalSettings = {
                    ColorMode = ColorMode.Color,
                    Orientation = Orientation.Portrait,
                    PaperSize = PaperKind.A4,
                },
                Objects = {
                    new ObjectSettings() {
                        HtmlContent = htmlContent,
                    },
                }
            };

            var pdf = converter.Convert(doc);

            string pdfPath = Path.Combine(Path.GetTempPath(), string.Format("{0}_output.pdf", orderNumber));
            File.WriteAllBytes(pdfPath, pdf);

            return pdfPath;
        }

        private static void PrintPdf(string pdfPath, string printerName)
        {
            using (var document = PdfDocument.Load(pdfPath))
            {
                using (var printDocument = document.CreatePrintDocument())
                {
                    printDocument.PrinterSettings = new PrinterSettings
                    {
                        PrinterName = printerName,
                        PrintFileName = pdfPath,
                    };

                    printDocument.DefaultPageSettings.PaperSize = new PaperSize("A4", 827, 1169);
                    printDocument.DefaultPageSettings.Color = true;

                    printDocument.Print();
                }
            }
        }
    }
}
