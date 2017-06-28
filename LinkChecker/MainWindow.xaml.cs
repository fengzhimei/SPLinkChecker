using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows.Threading;
using System.Linq;
using System.Diagnostics;
using System.Data;
using System.Windows;
using WinDoc = System.Windows.Documents;
using System.Windows.Media;
using System.Windows.Navigation;
using System.Collections;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using WordProcessing = DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.SharePoint.Client;
using OfficeOpenXml;
using HtmlAgilityPack;
using LumiSoft.Net.Mail;
using System.Net.Sockets;

namespace LinkChecker
{
    public partial class MainWindow : Window
    {
        private string siteURL = string.Empty;
        private string libriaryName = string.Empty;
        private Uri siteURI = null;
        private ClientContext spContext = null;
        private Microsoft.SharePoint.Client.List spLibrary = null;
        private HashSet<string> brokenLinksFound = new HashSet<string>();
        private HashSet<string> errorHost = new HashSet<string>();

        public MainWindow()
        {
            InitializeComponent();
            this.Loaded += MainWindow_Loaded;
        }

        void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            txtNote.Text = Environment.UserDomainName + "\\" + Environment.UserName + " must have permission to access the library.";
            spTimer.Visibility = Visibility.Collapsed;
            logSP.Visibility = Visibility.Collapsed;
        }
        
        private void btnCheck_Click(object sender, RoutedEventArgs e)
        {
            siteURL = txtSiteURL.Text.Trim();
            libriaryName = txtLibraryName.Text.Trim();

            string errorMsg = string.Empty;

            bool isInputValid = VerifySiteAndLibriary(siteURL, libriaryName, out errorMsg);

            if (!isInputValid)
            {
                MessageBox.Show(errorMsg, "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            //use thread to process the method. 
            Thread thread = new Thread(new ThreadStart(GetDcouments));
            thread.IsBackground = true;
            thread.Start();

            //start timer
            TimerStart();
        }

        private void GetDcouments()
        {
            Utility.ForceToUseNTLM();

            DataTable excelTable = new DataTable("LinkResults");
            DataColumn column0 = new DataColumn("File");
            DataColumn column1 = new DataColumn("Link Text");
            DataColumn column2 = new DataColumn("Link Address");
            excelTable.Columns.Add(column0);
            excelTable.Columns.Add(column1);
            excelTable.Columns.Add(column2);

            List<ListItem> itemsList = new List<ListItem>();
            int rowLimit = Settings.Default.RowLimit;
            int pageIndex = 0;

            CamlQuery spQuery = new CamlQuery();
            spQuery.ViewXml = "<View Scope='Recursive'>" + 
                                "<Query>" +
                                    "<Where>" +
                                        "<Or>" +
                                            "<Or>" +
                                                "<Or>" +
                                                    "<Eq>" +
                                                        "<FieldRef Name=\"DocIcon\" />" +
                                                        "<Value Type=\"Computed\">pptx</Value>" +
                                                    "</Eq>" +
                                                    "<Eq>" +
                                                        "<FieldRef Name=\"DocIcon\" />" +
                                                        "<Value Type=\"Computed\">docx</Value>" +
                                                    "</Eq>" +
                                                "</Or>" +
                                                "<Eq>" +
                                                    "<FieldRef Name=\"DocIcon\" />" +
                                                    "<Value Type=\"Computed\">mht</Value>" +
                                                "</Eq>" +
                                            "</Or>" +
                                            "<Eq>" +
                                                "<FieldRef Name=\"DocIcon\" />" +
                                                "<Value Type=\"Computed\">aspx</Value>" +
                                            "</Eq>" +
                                        "</Or>" +
                                    "</Where>" +
                                "</Query>" + 
                                "<RowLimit>" + rowLimit + "</RowLimit>" + 
                                "<ViewFields>" +
                                    "<FieldRef Name='File_x0020_Type' />" +
                                    "<FieldRef Name='FileLeafRef' />" +
                                    "<FieldRef Name='FileDirRef' />" +
                                    "<FieldRef Name='FileRef' />" + 
                                "</ViewFields>" +
                              "</View>";

            do
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    txtProgress.Inlines.Add(new WinDoc.Run() { Text = "Fetching " + (rowLimit * pageIndex + 1) + " to " + rowLimit * (pageIndex + 1) + " documents, please wait...", Foreground = Brushes.Green });
                    txtProgress.Inlines.Add(new WinDoc.LineBreak());
                    svProgress.ScrollToEnd();
                }), null);

                ListItemCollection listItems = spLibrary.GetItems(spQuery);
                spContext.Load(listItems);
                spContext.ExecuteQuery();

                itemsList.AddRange(listItems);

                spQuery.ListItemCollectionPosition = listItems.ListItemCollectionPosition;
                pageIndex++;

            } while (spQuery.ListItemCollectionPosition != null);

            int totalCount = itemsList.Count;
            this.Dispatcher.Invoke(new Action(() =>
            {
                txtProgress.Inlines.Add(new WinDoc.Run() { Text = "Finish fetching all " + totalCount + " documents, will extract links from these documents.", Foreground = Brushes.Green });
                txtProgress.Inlines.Add(new WinDoc.LineBreak());
                svProgress.ScrollToEnd();
            }), null);

            int fileIndex = 1;

            string siteRootURL = string.Empty;
            if (siteURI.Port == 80)
            {
                siteRootURL = siteURI.Scheme + "://" + siteURI.Host;
            }
            else
            {
                siteRootURL = siteURI.Scheme + "://" + siteURI.Host + ":" + siteURI.Port;
            }

            ParserFacade facade = new ParserFacade();

            foreach (ListItem item in itemsList)
            {
                if (item.FileSystemObjectType == FileSystemObjectType.File)
                {
                    string fileName = item.FieldValues["FileLeafRef"].ToString();
                    string fileUri = siteRootURL + (string)item.FieldValues["FileRef"];

                    this.Dispatcher.Invoke(new Action(() =>
                    {
                        txtProgress.Inlines.Add(new WinDoc.Run() { Text = fileIndex + "/" + totalCount + " -> " + fileName });
                        txtProgress.Inlines.Add(new WinDoc.LineBreak());
                        svProgress.ScrollToEnd();
                    }), null);

                    fileIndex++;

                    facade.ParseFile(fileUri);
                }
            }

            List<FileLink> results = facade.GetParseResult();
            int linkIndex = 1;
            int totalLinkCount = results.Count;

//#if DEBUG
//            string filePathAllResults = Environment.CurrentDirectory + "\\" + libriaryName + "_allResults.csv";

//            if (!System.IO.File.Exists(filePathAllResults))
//            {
//                System.IO.File.Create(filePathAllResults).Close();
//            }

//            using (System.IO.TextWriter writer = System.IO.File.CreateText(filePathAllResults))
//            {
//                for (int index = 0; index < totalLinkCount; index++)
//                {
//                    writer.WriteLine(string.Join(",", results[index].LinkAddress));
//                }
//            }
//#endif

            this.Dispatcher.Invoke(new Action(() =>
            {
                txtProgress.Inlines.Add(new WinDoc.Run() { Text = "There are " + totalLinkCount + " links to be valided, please wait...", Foreground = Brushes.Green });
                txtProgress.Inlines.Add(new WinDoc.LineBreak());
                svProgress.ScrollToEnd();
            }), null);

            foreach (FileLink fileLink in results)
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    txtProgress.Inlines.Add(new WinDoc.Run() { Text = linkIndex + "/" + totalLinkCount + " -> " + (fileLink.hasError ? "Invalid link" : fileLink.LinkAddress) });
                    txtProgress.Inlines.Add(new WinDoc.LineBreak());
                    svProgress.ScrollToEnd();
                }), null);

                linkIndex++;

                if (brokenLinksFound.Contains(fileLink.LinkAddress))
                {
                    CreateExcelRow(fileLink.ParentFileUrl, fileLink.LinkText, fileLink.LinkAddress, excelTable);
                }
                else
                {
                    if (!UrlIsValid(fileLink))
                    {
                        brokenLinksFound.Add(fileLink.LinkAddress);
                        CreateExcelRow(fileLink.ParentFileUrl, fileLink.LinkText, fileLink.LinkAddress, excelTable);
                    }
                }
            }

            this.Dispatcher.Invoke(new Action(() =>
            {
                txtProgress.Inlines.Add(new WinDoc.Run() { Text = "Finish validing links and writing results into excel, please waiting...", Foreground = Brushes.Green });
                txtProgress.Inlines.Add(new WinDoc.LineBreak());
                svProgress.ScrollToEnd();
            }), null);

            try
            {
                string logPath = Environment.CurrentDirectory + "\\" + libriaryName + "_Results_" + System.DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx";

                if (System.IO.File.Exists(logPath))
                {
                    System.IO.File.Delete(logPath);
                }

                FileInfo logFile = new FileInfo(logPath);

                using (ExcelPackage pck = new ExcelPackage(logFile))
                {
                    ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Broken Links");
                    ws.Cells["A1"].LoadFromDataTable(excelTable, true);
                    pck.Save();
                }
            
                this.Dispatcher.Invoke(new Action(() =>
                {
                    txtProgress.Inlines.Add(new WinDoc.Run() { Text = "Finished, please open below excel report to view results.", Foreground = Brushes.Green });
                    txtProgress.Inlines.Add(new WinDoc.LineBreak());
                    svProgress.ScrollToEnd();

                    logSP.Visibility = Visibility.Visible;
                    WinDoc.Hyperlink link = new WinDoc.Hyperlink(new WinDoc.Run(logPath));
                    link.NavigateUri = new Uri(logPath);
                    link.RequestNavigate += delegate(object sender, RequestNavigateEventArgs e) 
                    {
                        Process.Start(new ProcessStartInfo(e.Uri.AbsoluteUri));
                        e.Handled = true;
                    };
                    LogLink.Inlines.Add(link);

                }), null);
            }
            catch
            {
                this.Dispatcher.Invoke(new Action(() =>
                {
                    txtProgress.Inlines.Add(new WinDoc.Run() { Text = "Writing results to excel file failed.", Foreground = Brushes.Red });
                    txtProgress.Inlines.Add(new WinDoc.LineBreak());
                    svProgress.ScrollToEnd();
                }), null);
            }

            TimerStop();

        }

        #region timer
        DispatcherTimer myTimer = null;

        int spendTime = 0;

        void myTimer_Tick(object sender, EventArgs e)
        {
            spendTime++;
            int m = spendTime / 60;
            int h = m / 60;
            m = m % 60;
            int s = spendTime % 60;
            txtDuration.Text = (h.ToString().Length > 1 ? h.ToString() : h.ToString().PadLeft(2, '0')) + ":" + (m.ToString().Length > 1 ? m.ToString() : m.ToString().PadLeft(2, '0')) + ":" + (s.ToString().Length > 1 ? s.ToString() : s.ToString().PadLeft(2, '0'));
        }

        void TimerStart()
        {
            btnCheck.IsEnabled = false;
            spTimer.Visibility = Visibility.Visible;
            spendTime = 0;
            txtDuration.Text = "00:00:00";
            txtProgress.Text = string.Empty;
            myTimer = new DispatcherTimer();
            myTimer.Interval = new TimeSpan(0, 0, 0, 1);
            myTimer.Tick += new EventHandler(myTimer_Tick);
            myTimer.Start();
        }

        void TimerStop()
        {
            this.Dispatcher.Invoke(new Action(() =>
            {
                btnCheck.IsEnabled = true;
            }), null);

            if (myTimer != null)
            {
                myTimer.Stop();

            }
        }
        #endregion timer

        private bool VerifySiteAndLibriary(string siteURL, string libriaryName, out string returnMsg)
        {
            returnMsg = string.Empty;

            if (string.IsNullOrEmpty(siteURL)
                || !(siteURL.StartsWith("http://", StringComparison.InvariantCultureIgnoreCase) || siteURL.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase))
                || !Uri.TryCreate(siteURL, UriKind.Absolute, out siteURI))
            {
                returnMsg = "Please input a valid URL for the site (start with http:// or https://).";
                return false;
            }

            if (string.IsNullOrEmpty(libriaryName))
            {
                returnMsg = "Please input a Name for the library.";
                return false;
            }

            try
            {
                spContext = new ClientContext(siteURL);
                spLibrary = spContext.Web.Lists.GetByTitle(libriaryName);
                spContext.Load(spLibrary);
                spContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                returnMsg = ex.Message;
                return false;
            }

            return true;
        }

        private bool UrlIsValid(FileLink fileLink)
        {
            if (fileLink.hasError)
                return false;

            string formatedURL = fileLink.LinkAddress;
            Uri uri = null;

            Uri.TryCreate(formatedURL, UriKind.Absolute, out uri);

            if (uri == null)
            {
                if (formatedURL.StartsWith("/"))
                {
                    Uri originalURI = null;
                    Uri.TryCreate(fileLink.ParentFileUrl, UriKind.Absolute, out originalURI);
                    if (originalURI != null)
                    {
                        formatedURL = originalURI.Scheme + "://" + originalURI.Host + ":" + originalURI.Port + formatedURL;
                    }
                }
            }

            Uri.TryCreate(formatedURL, UriKind.Absolute, out uri);

            if (uri == null)
                return false;

            try
            {
                if (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)
                {
                    if (errorHost.Contains(uri.Host))
                        return false;

                    if (uri.Scheme == Uri.UriSchemeHttps)
                    {
                        ServicePointManager.ServerCertificateValidationCallback += (sender, certificate, chain, sslPolicyErrors) => true;
                    }

                    try
                    {
                        HttpWebRequest hwReq = HttpWebRequest.Create(uri) as HttpWebRequest;
                        hwReq.AutomaticDecompression = DecompressionMethods.GZip;
                        hwReq.UseDefaultCredentials = true;
                        hwReq.UserAgent = "uKnowlive Link Checker Tool";
                        hwReq.Timeout = 15000;
                        hwReq.Method = "HEAD";
                        hwReq.KeepAlive = false;

                        using (HttpWebResponse hwRes = hwReq.GetResponse() as HttpWebResponse)
                        {
                            return (hwRes == null ? false : hwRes.StatusCode == HttpStatusCode.OK);
                        }
                    }
                    catch (WebException wex)
                    {
                        if (wex.Status == WebExceptionStatus.NameResolutionFailure)
                        {
                            if (errorHost.Contains(uri.Host) == false)
                                errorHost.Add(uri.Host);
                        }
                        else if (wex.Status == WebExceptionStatus.ProtocolError)
                        {
                            if (wex.Response != null)
                            {
                                HttpWebResponse response = wex.Response as HttpWebResponse;
                                if (response.StatusCode == HttpStatusCode.NotFound)
                                {
                                    try
                                    {
                                        HttpWebRequest hwReqNew = HttpWebRequest.Create(uri) as HttpWebRequest;
                                        hwReqNew.AutomaticDecompression = DecompressionMethods.GZip;
                                        hwReqNew.UseDefaultCredentials = true;
                                        hwReqNew.UserAgent = "uKnowlive Link Checker Tool";
                                        hwReqNew.Timeout = 15000;
                                        hwReqNew.Method = "GET";
                                        hwReqNew.KeepAlive = false;

                                        using (HttpWebResponse hwResNew = hwReqNew.GetResponse() as HttpWebResponse)
                                        {
                                            try
                                            {
                                                hwReqNew.Abort();
                                            }
                                            catch{}
                                            return (hwResNew == null ? false : hwResNew.StatusCode == HttpStatusCode.OK);
                                        }
                                    }
                                    catch{}
                                }
                                //else if (response.StatusCode == HttpStatusCode.Unauthorized)
                                //{
                                //    try
                                //    {
                                //        CredentialCache credCache = new CredentialCache();
                                //        credCache.Add(uri, "NTLM", CredentialCache.DefaultNetworkCredentials);

                                //        HttpWebRequest hwReqNew = HttpWebRequest.Create(uri) as HttpWebRequest;
                                //        hwReqNew.AutomaticDecompression = DecompressionMethods.GZip;
                                //        hwReqNew.Credentials = credCache;
                                //        hwReqNew.UserAgent = "uKnowlive Link Checker Tool";
                                //        hwReqNew.Timeout = 15000;
                                //        hwReqNew.Method = "HEAD";
                                //        hwReqNew.KeepAlive = false;

                                //        using (HttpWebResponse hwResNew = hwReqNew.GetResponse() as HttpWebResponse)
                                //        {
                                //            return (hwResNew == null ? false : hwResNew.StatusCode == HttpStatusCode.OK);
                                //        }
                                //    }
                                //    catch { }
                                //}
                            }
                        }
                    }
                }
                else if (uri.Scheme == Uri.UriSchemeFtp)
                {
                    FtpWebRequest ftpWebRequest = FtpWebRequest.Create(uri) as FtpWebRequest;
                    ftpWebRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                    ftpWebRequest.KeepAlive = false;

                    using (FtpWebResponse ftpRes = ftpWebRequest.GetResponse() as FtpWebResponse)
                    {
                        return (ftpRes != null);
                    }
                }
                else if (uri.Scheme == Uri.UriSchemeFile)
                {
                    if (System.IO.Path.HasExtension(uri.LocalPath))
                    {
                        FileInfo fi = new FileInfo(uri.LocalPath);
                        return fi.Exists;
                    }
                    else if (uri.Segments.Length > 1)
                    {
                        DirectoryInfo di = new DirectoryInfo(uri.LocalPath);
                        return di.Exists;
                    }
                    else
                    {
                        IPHostEntry iph = Dns.GetHostEntry(uri.Host);
                        if (iph != null && iph.AddressList.Length > 0)
                        {
                            return true;
                        }
                    }
                }
                else if (uri.Scheme == Uri.UriSchemeMailto)
                {
                    return true;
                }
            }
            catch{}

            return false;
        }

        private void CreateExcelRow(string fileUrl, string linkText, string linkUrl, DataTable excelTable)
        {
            DataRow row = excelTable.NewRow();
            row[0] = fileUrl;
            row[1] = linkText;
            row[2] = linkUrl;
            excelTable.Rows.Add(row);
        }
    }
}
