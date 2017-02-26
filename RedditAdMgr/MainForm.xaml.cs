using MahApps.Metro.Controls;
using System.Net;
using System.Windows;
using System.IO;
using OfficeOpenXml;
using Microsoft.Win32;
using System.Threading.Tasks;
using System.Data;
using System.Collections.Generic;
using RedditAdMgr.Model;
using System.Linq;
using System;
using System.Threading;
using HtmlAgilityPack;
using System.Text;

namespace RedditAdMgr
{
    /// <summary>
    /// Interaction logic for MainForm.xaml
    /// </summary>
    public partial class MainForm : MetroWindow
    {
        internal CookieContainer Cookies { get; set; }
        private List<Advertisement> ads { get; set; }
        private List<Campaign> campaigns { get; set; }

        public MainForm()
        {
            InitializeComponent();
            BeginCreationButton.IsEnabled = false;
            GeneralProgressBar.Visibility = Visibility.Hidden;
        }

        private void ImageExplorerButton_Click(object sender, RoutedEventArgs e)
        {
            using (System.Windows.Forms.FolderBrowserDialog folderBrowser = new System.Windows.Forms.FolderBrowserDialog())
            {
                if (folderBrowser.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    ImagePathTextBox.Text = folderBrowser.SelectedPath;
            }
        }

        private void ExcelExplorerButton_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Excel files (*.xls; *.xlsx; *.xlsm) | *.xls; *.xlsx; *.xlsm";

            if (openDialog.ShowDialog() == true)
                ExcelPathTextBox.Text = openDialog.FileName;
        }

        private void ImportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ExcelPathTextBox.Text) && !string.IsNullOrEmpty(ImagePathTextBox.Text))
            {
                var FileInfo = new FileInfo(ExcelPathTextBox.Text);
                var path = ExcelPathTextBox.Text;
                bool error = false;
                string errorMsg = string.Empty;

                GeneralProgressBar.Visibility = Visibility.Visible;
                GeneralProgressBar.IsIndeterminate = true;

                Log("INFO", string.Format("Loading file {0}",FileInfo.Name));

                Task readData = Task.Factory.StartNew(() => ReadExcelFile(path));

                try
                {
                    readData.Wait();
                }
                catch(AggregateException ae)
                {
                    ae.Handle((x) =>
                    {
                        errorMsg = x.Message + " | " + x.StackTrace;
                        error = true;

                        return error;
                    });
                }

                GeneralProgressBar.Visibility = Visibility.Hidden;
                Log("INFO", "File loaded");

                if (error)
                {
                    Log("ERROR", errorMsg);
                }
                else
                {
                    Log("INFO", string.Format("ToDo: Creating {0} ads", ads.Count));
                    Log("INFO", string.Format("ToDo: Creating {0} campaigns", campaigns.Count));
                }

                BeginCreationButton.IsEnabled = true;
            }
            else
                MessageBox.Show("Please select an image and Excel path before proceeding!", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void Log(string type, string msg)
        {
            string status = string.Empty;

            switch(type)
            {
                case "INFO":

                    status += string.Format("[INFO] {0} - {1}", DateTime.Now.ToString(), msg);

                    break;

                case "ERROR":

                    status += string.Format("[ERROR] {0} - {1}", DateTime.Now.ToString(), msg);

                    break;

                case "WARNING":

                    status += string.Format("[WARNING] {0} - {1}", DateTime.Now.ToString(), msg);

                    break;
            }

            StatusTextBlock.Text += status + Environment.NewLine;
            StatusTextBlock.ScrollToEnd();

            using (var writer = new StreamWriter("Log.txt", true))
            {
                writer.Write(status + Environment.NewLine);
            }
        }

        private void ReadExcelFile(string path)
        {
            var adsDT = GetDataTableFromExcel(path, "advertisements", true);
            ads = new List<Advertisement>();
            campaigns = new List<Campaign>();

            ads = adsDT.AsEnumerable().Select(row =>
                new Advertisement
                {
                    AdvertisementNumber = Convert.ToInt32(row.Field<string>("ADVERTISEMENT NUMBER")),
                    ThumbnailName = row.Field<string>("THUMBNAIL NAME"),
                    Title = row.Field<string>("TITLE"),
                    Url = row.Field<string>("URL"),
                    DisableComments = (row.Field<string>("OPTION_DISABLECOMMENTS") == "1") ? true : false,
                    SendComments = (row.Field<string>("OPTION_SENDCOMMENTS") == "1") ? true : false
                }).ToList();

            var campsDT = GetDataTableFromExcel(path, "campaigns", true);

            campaigns = campsDT.AsEnumerable().Select(row =>
                new Campaign
                {
                    Advertisement = ads.Where(ad => ad.AdvertisementNumber == Convert.ToInt32(row.Field<string>("WHICH ADVERTISEMENT?"))).First(),
                    Target = row.Field<string>("TARGET"),
                    TargetDetail = row.Field<string>("TARGET_DETAIL"),
                    Location = row.Field<string>("LOCATION"),
                    Location2 = row.Field<string>("LOCATION_2"),
                    Platform = row.Field<string>("PLATFORM"),
                    Budget = Convert.ToDecimal(row.Field<string>("BUDGET")),
                    BudgetOptionDeliverFast = (row.Field<string>("BUDGET_OPTION_DELIVERFAST") == "1") ? true : false,
                    Start = Convert.ToDateTime(row.Field<string>("START")),
                    End = Convert.ToDateTime(row.Field<string>("END")),
                    OptionExtend = (row.Field<string>("OPTION_EXTEND") == "1") ? true : false,
                    PricingCPM = Convert.ToDecimal(row.Field<string>("PRICINGCPM"))
                }).ToList();
        }

        public DataTable GetDataTableFromExcel(string path, string worksheetName, bool hasHeader = true)
        {
            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                ExcelWorksheet ws = null;

                ws = pck.Workbook.Worksheets[worksheetName];
                
                DataTable tbl = new DataTable();

                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }

                var startRow = hasHeader ? 2 : 1;

                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }

                return tbl;
            }
        }

        private async void BeginCreationButton_Click(object sender, RoutedEventArgs e)
        {
            Log("INFO", "Beggining ad creation");
            GeneralProgressBar.Visibility = Visibility.Visible;
            foreach(var ad in ads)
            {
                string adUh = string.Empty;
                string errorMsg = string.Empty;
                bool error = false;

                //Task<string> uhFetch = Task<string>.Factory.StartNew(() => GetUhForAdvert());
                Task task = Task.Factory.StartNew(() =>
                {
                    HttpWebRequest newPromoRequest = WebRequest.Create("https://www.reddit.com/promoted/new_promo/") as HttpWebRequest;
                    newPromoRequest.CookieContainer = Cookies;
                    newPromoRequest.Method = "GET";
                    newPromoRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                    string uh = string.Empty;

                    HttpWebResponse response = (HttpWebResponse)newPromoRequest.GetResponse();

                    if (response.StatusCode == HttpStatusCode.OK)
                    {
                        using (Stream s = response.GetResponseStream())
                        {
                            using (StreamReader sr = new StreamReader(s, Encoding.GetEncoding(response.CharacterSet)))
                            {
                                HtmlDocument doc = new HtmlDocument();

                                doc.Load(sr);

                                adUh = doc.DocumentNode.SelectNodes("//form")[0].SelectNodes("//input")[0].Attributes[2].Value;
                            }
                        }
                    }

                    string postString = string.Format("uh={0}&id=%23promo-form&title={1}&kind=link&url={2}&thing_id=&text=&renderstyle=html", adUh, ad.Title, ad.Url);
                    string adUrl = "https://www.reddit.com/api/create_promo";

                    if (ad.SendComments)
                        postString += "&sendreplies=on";

                    if (ad.DisableComments)
                        postString += "&disable_comments=on";

                    HttpWebRequest promoRequest = WebRequest.Create(adUrl) as HttpWebRequest;
                    promoRequest.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                    promoRequest.Method = "POST";
                    promoRequest.CookieContainer = Cookies;
                    promoRequest.Accept = "application/json, text/javascript, */*; q=0.01";
                    promoRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
                    promoRequest.UserAgent = "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36";
                    promoRequest.Referer = "https://www.reddit.com/";
                    
                    WebHeaderCollection customHeaders = promoRequest.Headers;

                    customHeaders.Add("accept-language", "en;q=0.4");
                    customHeaders.Add("origin", "https://www.reddit.com");
                    customHeaders.Add("x-requested-with", "XMLHttpRequest");

                    byte[] bytes = Encoding.ASCII.GetBytes(postString);

                    promoRequest.ContentLength = bytes.Length;

                    using (Stream os = promoRequest.GetRequestStream())
                    {
                        os.Write(bytes, 0, bytes.Length);
                    }

                    //"your message has been delivered"

                    HttpWebResponse promoResponse = promoRequest.GetResponse() as HttpWebResponse;


                    if (promoResponse.StatusCode == HttpStatusCode.OK)
                    {
                        using (Stream s = promoResponse.GetResponseStream())
                        {
                            using (StreamReader sr = new StreamReader(s, Encoding.GetEncoding(promoResponse.CharacterSet)))
                            {
                                if (sr.ReadToEnd().Contains("your message has been delivered"))
                                    Log("INFO", string.Format("Ad #{0} successfully created", ad.AdvertisementNumber));
                            }
                        }
                    }
                });

                try
                {
                    await task;
                }
                catch (AggregateException ae)
                {
                    ae.Handle((x) =>
                    {
                        errorMsg = x.Message + " | " + x.StackTrace;
                        error = true;

                        return error;
                    });
                }

                if (error)
                {
                    Log("ERROR", errorMsg);
                }

                Log("INFO", adUh);
                //adUh = uhFetch.Result;

            //    //uh = 5xladjq2rzdb809b37a0dd4580c3afbb17284b81106662be19 & id =% 23promo - form & title = alidsuhalsisdlasd & kind = link & url = test.com & thing_id = &text = &disable_comments = on & sendreplies = on & renderstyle = html
            }
        }

        private string GetUhForAdvert()
        {
            HttpWebRequest newPromoRequest = WebRequest.Create("https://www.reddit.com/promoted/new_promo/") as HttpWebRequest;
            newPromoRequest.CookieContainer = Cookies;
            newPromoRequest.Method = "GET";
            newPromoRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            string uh = string.Empty;

            HttpWebResponse newMsgResponse = (HttpWebResponse)newPromoRequest.GetResponse();

            if (newMsgResponse.StatusCode == HttpStatusCode.OK)
            {
                using (Stream s = newMsgResponse.GetResponseStream())
                {
                    using (StreamReader sr = new StreamReader(s, Encoding.GetEncoding(newMsgResponse.CharacterSet)))
                    {
                        HtmlDocument doc = new HtmlDocument();

                        doc.Load(sr);

                        uh = doc.DocumentNode.SelectNodes("//form")[0].SelectNodes("//input")[0].Attributes[2].Value;
                    }
                }
            }

            return uh;
        }
}
}
