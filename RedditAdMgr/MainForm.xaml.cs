using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json;
using OfficeOpenXml;
using RedditAdMgr.Model;
using RedditAdMgr.Utils;
using Microsoft.WindowsAPICodePack.Dialogs;
using HtmlDocument = HtmlAgilityPack.HtmlDocument;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

// ReSharper disable AssignNullToNotNullAttribute
namespace RedditAdMgr
{
    /// <summary>
    ///     Interaction logic for MainForm.xaml
    /// </summary>
    [SuppressMessage("ReSharper", "UseFormatSpecifierInFormatString")]
    public partial class MainForm
    {
        internal CookieContainer Cookies { get; set; }
        private List<Advertisement> Ads { get; set; }
        private List<Campaign> Campaigns { get; set; }

        public MainForm()
        {
            InitializeComponent();

            BeginCreationButton.IsEnabled = false;
            GeneralProgressBar.Visibility = Visibility.Hidden;
            DelayPicker.Value = 1000;
        }

        /// <summary>
        /// Open a folder picker dialog to get images' directory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImageExplorerButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog {IsFolderPicker = true};

            if (dialog.ShowDialog() ==CommonFileDialogResult.Ok)
                ImagePathTextBox.Text = dialog.FileName;
        }

        /// <summary>
        /// Open a file picker dialog to get data Excel path
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ExcelExplorerButton_Click(object sender, RoutedEventArgs e)
        {
            var openDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xls; *.xlsx; *.xlsm) | *.xls; *.xlsx; *.xlsm"
            };

            if (openDialog.ShowDialog() == true)
                ExcelPathTextBox.Text = openDialog.FileName;
        }

        /// <summary>
        /// Import Excel data into memory
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ImportExcelButton_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(ExcelPathTextBox.Text) && !string.IsNullOrEmpty(ImagePathTextBox.Text))
            {
                var fileInfo = new FileInfo(ExcelPathTextBox.Text);
                var path = ExcelPathTextBox.Text;
                var error = false;
                var errorMsg = string.Empty;

                GeneralProgressBar.Visibility = Visibility.Visible;
                GeneralProgressBar.IsIndeterminate = true;

                Log("INFO", $"Loading file {fileInfo.Name}");

                var readData = Task.Factory.StartNew(() => ReadExcelFile(path));

                try
                {
                    readData.Wait();
                }
                catch (AggregateException ae)
                {
                    ae.Handle(x =>
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
                    Log("INFO", $"ToDo: Creating {Ads.Count} ads");
                    Log("INFO", $"ToDo: Creating {Campaigns.Count} campaigns");
                }

                BeginCreationButton.IsEnabled = true;
            }
            else
            {
                MessageBox.Show("Please select an image and Excel path before proceeding!", "ERROR", MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Register events to MainWindow status textblock and Log file
        /// </summary>
        /// <param name="type">Log message type (INFO, WARNING, ERROR)</param>
        /// <param name="msg">Message to display</param>
        private void Log(string type, string msg)
        {
            var status = string.Empty;

            switch (type)
            {
                case "INFO":

                    status += $"[INFO] {DateTime.Now.ToString(CultureInfo.CurrentCulture)} - {msg}";

                    break;

                case "ERROR":

                    status += $"[ERROR] {DateTime.Now.ToString(CultureInfo.CurrentCulture)} - {msg}";

                    break;

                case "WARNING":

                    status += $"[WARNING] {DateTime.Now.ToString(CultureInfo.CurrentCulture)} - {msg}";

                    break;
            }

            StatusTextBlock.Text += status + Environment.NewLine;
            StatusTextBlock.ScrollToEnd();

            using (var writer = new StreamWriter("Log.txt", true))
            {
                writer.Write(status + Environment.NewLine);
            }
        }

        /// <summary>
        /// Read Excel file into memory filling Ads and Campaigns lists
        /// </summary>
        /// <param name="path">Excel file path</param>
        private void ReadExcelFile(string path)
        {
            var adsDt = GetDataTableFromExcel(path, "advertisements");
            Ads = new List<Advertisement>();
            Campaigns = new List<Campaign>();

            Ads = adsDt.AsEnumerable().Select(row =>
                new Advertisement
                {
                    AdvertisementNumber = Convert.ToInt32(row.Field<string>("ADVERTISEMENT NUMBER")),
                    ThumbnailName = row.Field<string>("THUMBNAIL NAME"),
                    Title = row.Field<string>("TITLE"),
                    Url = row.Field<string>("URL"),
                    DisableComments = row.Field<string>("OPTION_DISABLECOMMENTS") == "1",
                    SendComments = row.Field<string>("OPTION_SENDCOMMENTS") == "1"
                }).ToList();

            var campsDt = GetDataTableFromExcel(path, "campaigns");

            Campaigns = campsDt.AsEnumerable().Select(row =>
                new Campaign
                {
                    Advertisement =
                        Ads.First(
                            ad => ad.AdvertisementNumber == Convert.ToInt32(row.Field<string>("WHICH ADVERTISEMENT?"))),
                    Target = row.Field<string>("TARGET"),
                    TargetDetail = row.Field<string>("TARGET_DETAIL"),
                    Location = row.Field<string>("LOCATION"),
                    Location2 = row.Field<string>("LOCATION_2"),
                    Platform = row.Field<string>("PLATFORM"),
                    Budget = Convert.ToDecimal(row.Field<string>("BUDGET")),
                    BudgetOptionDeliverFast = row.Field<string>("BUDGET_OPTION_DELIVERFAST") == "1",
                    Start = Convert.ToDateTime(row.Field<string>("START")),
                    End = Convert.ToDateTime(row.Field<string>("END")),
                    OptionExtend = row.Field<string>("OPTION_EXTEND") == "1",
                    PricingCpm = Convert.ToDecimal(row.Field<string>("PRICINGCPM"))
                }).ToList();
        }

        /// <summary>
        /// Read an Excel file, fill content into a DataTable
        /// </summary>
        /// <param name="path">Excel path</param>
        /// <param name="worksheetName">Worksheet where data is placed</param>
        /// <param name="hasHeader">Determine if worksheet has header or not</param>
        /// <returns></returns>
        public DataTable GetDataTableFromExcel(string path, string worksheetName, bool hasHeader = true)
        {
            using (var pck = new ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }

                var ws = pck.Workbook.Worksheets[worksheetName];

                var tbl = new DataTable();

                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");

                var startRow = hasHeader ? 2 : 1;

                for (var rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    var row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                        row[cell.Start.Column - 1] = cell.Text;
                }

                return tbl;
            }
        }

        /// <summary>
        /// Handle ads and campaigns processes
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private async void BeginCreationButton_Click(object sender, RoutedEventArgs e)
        {
            Log("INFO", "Beggining ad creation");
            GeneralProgressBar.Visibility = Visibility.Visible;

            foreach (var ad in Ads)
            {
                var errorMsg = string.Empty;
                var error = false;
                var imgPathText = ImagePathTextBox.Text;
                var result = new RedditAdJson();

                try
                {
                    var task = Task.Factory.StartNew(() =>
                    {
                        result = StartAdvertisementProcess(ad, imgPathText);
                    });
                    
                    await task;
                }
                catch (AggregateException ae)
                {
                    ae.Handle(x =>
                    {
                        if (x is FileNotFoundException)
                        {
                            errorMsg = x.Message;
                            error = true;
                        }
                        else
                        {
                            errorMsg = x.Message + " | " + x.StackTrace;
                            error = true;
                        }

                        return error;
                    });
                }
                catch (Exception ex)
                {
                    errorMsg = ex.Message + " | " + ex.StackTrace;
                    error = true;
                }

                if (error)
                {
                    Log("ERROR", errorMsg);
                }
                else
                {
                    if (result.Success)
                    {
                        ad.RedditAdId = result.Jquery[16][3].ToString().Split('/')[5].Remove(6);
                        Log("INFO", $"Ad #{ad.AdvertisementNumber} successfully created ({ad.RedditAdId})");
                    }
                }

                await Task.Delay(Convert.ToInt32(DelayPicker.Value));
            }

            Log("INFO", "Finished creating ads!");

            // Update Reddit ad ID on campaigns list
            foreach (var camp in Campaigns)
                camp.Advertisement.RedditAdId =
                    Ads.Where(x => x.AdvertisementNumber == camp.Advertisement.AdvertisementNumber)
                        .Select(x => x.RedditAdId)
                        .First();


            Log("INFO", "Beggining campaign creation");

            foreach (var camp in Campaigns)
            {
                var errorMsg = string.Empty;
                var error = false;
                var result = new RedditAdJson();

                try
                {
                    var task = Task.Factory.StartNew(() =>
                    {
                        result = StartCampaignProcess(camp);
                    });
                    
                    await task;
                }
                catch (AggregateException ae)
                {
                    ae.Handle(x =>
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
                else
                {
                    if (result.Success)
                    {
                        Log("INFO",
                            $"Campaign for ad ID {camp.Advertisement.RedditAdId} successfully created");
                    }
                    else
                    {
                        var msg = result.Jquery[14][3];
                        Log("ERROR",
                            $"Error creating campaign for ad ID {camp.Advertisement.RedditAdId}:{msg.ToString().Replace(Environment.NewLine, string.Empty).Replace("[", string.Empty).Replace("]", string.Empty).Replace("\"", string.Empty)}");
                    }
                }

                await Task.Delay(Convert.ToInt32(DelayPicker.Value));
            }

            GeneralProgressBar.Visibility = Visibility.Hidden;
            Log("INFO", "Completed");
        }

        /// <summary>
        /// Method to handle campaign creation
        /// </summary>
        /// <param name="camp">Campaign to be created</param>
        /// <returns>A status object stating whether the creation failed or succeeded</returns>
        private RedditAdJson StartCampaignProcess(Campaign camp)
        {
            var campaignUh = GetCampaignUh(camp);

            return CreateCampaign(campaignUh, camp);
        }

        /// <summary>
        /// Create a campaign for the given parameters
        /// </summary>
        /// <param name="campaignUh">Reddit's campaign creation page modhash</param>
        /// <param name="camp">Campaign to be created</param>
        /// <returns>A status object stating whether the creation failed or succeeded</returns>
        private RedditAdJson CreateCampaign(string campaignUh, Campaign camp)
        {
            var result = new RedditAdJson();

            string postString =
                $"link_id36={camp.Advertisement.RedditAdId}&targeting=subreddit&sr=&selected_sr_names={camp.TargetDetail}&country={camp.Location}&region={camp.Location2}&metro=&mobile_os=&platform={camp.Platform}&undefined={camp.Start.AddDays(3).ToString("MM/dd/yyyy")}&total_budget_dollars={camp.Budget}&impressions={camp.Budget / 200 * 1000 * 100}&startdate={camp.Start.ToString("MM/dd/yyyy")}&enddate={camp.End.ToString("MM/dd/yyyy")}&cost_basis=cpm&bid_dollars={string.Format(CultureInfo.GetCultureInfo("en-US"), "{0: 0.##}", camp.PricingCpm)}&is_new=true&campaign_id36=&campaign_name=&id=%23campaign&uh={campaignUh}&renderstyle=html";

            var adUrl = "https://www.reddit.com/api/edit_campaign";

            if (camp.BudgetOptionDeliverFast)
                postString += "&no_daily_budget=on";

            if (camp.OptionExtend)
                postString += "&auto_extend=on";

            var campaignRequest = WebRequest.Create(adUrl) as HttpWebRequest;
            if (campaignRequest != null)
            {
                campaignRequest.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                campaignRequest.Method = "POST";
                campaignRequest.CookieContainer = Cookies;
                campaignRequest.Accept = "application/json, text/javascript, */*; q=0.01";
                campaignRequest.AutomaticDecompression = DecompressionMethods.GZip |
                                                         DecompressionMethods.Deflate;
                campaignRequest.UserAgent =
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36";
                campaignRequest.Referer = "https://www.reddit.com/";

                var customHeaders = campaignRequest.Headers;

                customHeaders.Add("accept-language", "en;q=0.4");
                customHeaders.Add("origin", "https://www.reddit.com");
                customHeaders.Add("x-requested-with", "XMLHttpRequest");

                var bytes = Encoding.ASCII.GetBytes(postString);

                campaignRequest.ContentLength = bytes.Length;

                using (var os = campaignRequest.GetRequestStream())
                {
                    os.Write(bytes, 0, bytes.Length);
                }

                var campaignResponse = campaignRequest.GetResponse() as HttpWebResponse;


                if (campaignResponse != null && campaignResponse.StatusCode == HttpStatusCode.OK)
                    using (var s = campaignResponse.GetResponseStream())
                    {
                        using (
                            // ReSharper disable once AssignNullToNotNullAttribute
                            var sr = new StreamReader(s, Encoding.GetEncoding(campaignResponse.CharacterSet))
                        )
                        {
                            var jsonString = sr.ReadToEnd();
                            result = JsonConvert.DeserializeObject<RedditAdJson>(jsonString);
                        }
                    }
            }

            return result;
        }

        /// <summary>
        /// Get a Reddit campaign creation page modhash to create a new campaign
        /// </summary>
        /// <param name="camp">campaign to be created</param>
        /// <returns>Reddit's campaign creation page modhash</returns>
        private string GetCampaignUh(Campaign camp)
        {
            var campaignUh = string.Empty;

            var newPromoRequest = WebRequest.Create(
                                $"https://www.reddit.com/promoted/edit_promo/{camp.Advertisement.RedditAdId}") as
                            HttpWebRequest;
            if (newPromoRequest != null)
            {
                newPromoRequest.CookieContainer = Cookies;
                newPromoRequest.Method = "GET";
                newPromoRequest.Accept =
                    "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";

                var response = (HttpWebResponse)newPromoRequest.GetResponse();

                if (response.StatusCode == HttpStatusCode.OK)
                    using (var s = response.GetResponseStream())
                    {
                        // ReSharper disable once AssignNullToNotNullAttribute
                        using (var sr = new StreamReader(s, Encoding.GetEncoding(response.CharacterSet)))
                        {
                            var doc = new HtmlDocument();

                            doc.Load(sr);

                            campaignUh =
                                doc.DocumentNode.SelectNodes("//form")[0].SelectNodes("//input")[0]
                                    .Attributes[2].Value;
                        }
                    }
            }

            return campaignUh;
        }

        /// <summary>
        /// Method to handle advertisement creation
        /// </summary>
        /// <param name="ad">Advertisement to be created</param>
        /// <param name="imgPathText">Images' path</param>
        /// <returns></returns>
        private RedditAdJson StartAdvertisementProcess(Advertisement ad, string imgPathText)
        {
            var adUh = GetAdUh();

            var adS3Data = GetAdS3Data(ad, adUh);

            UploadImage(adS3Data, imgPathText, ad);

            var result = CreateAdvertisement(adUh, ad);

            return result;
        }

        /// <summary>
        /// Create an advertisement for the given parameters
        /// </summary>
        /// <param name="adUh">Reddit's advertisement creation page modhash</param>
        /// <param name="ad">Advertisement to be created</param>
        /// <returns>A status object stating whether the creation failed or succeeded</returns>
        private RedditAdJson CreateAdvertisement(string adUh, Advertisement ad)
        {
            string postString =
                            $"uh={adUh}&id=%23promo-form&title={ad.Title}&kind=link&url={ad.Url}&thing_id=&text=&renderstyle=html";
            var adUrl = "https://www.reddit.com/api/create_promo";
            var result = new RedditAdJson();

            if (ad.SendComments)
                postString += "&sendreplies=on";

            if (ad.DisableComments)
                postString += "&disable_comments=on";

            var promoRequest = WebRequest.Create(adUrl) as HttpWebRequest;

            if (promoRequest == null) throw new ArgumentNullException(nameof(promoRequest));

            promoRequest.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
            promoRequest.Method = "POST";
            promoRequest.CookieContainer = Cookies;
            promoRequest.Accept = "application/json, text/javascript, */*; q=0.01";
            promoRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            promoRequest.UserAgent =
                "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36";
            promoRequest.Referer = "https://www.reddit.com/";
            promoRequest.Timeout = 15000;

            var customHeaders = promoRequest.Headers;

            customHeaders.Add("accept-language", "en;q=0.4");
            customHeaders.Add("origin", "https://www.reddit.com");
            customHeaders.Add("x-requested-with", "XMLHttpRequest");

            var bytes = Encoding.ASCII.GetBytes(postString);

            promoRequest.ContentLength = bytes.Length;

            using (var os = promoRequest.GetRequestStream())
            {
                os.Write(bytes, 0, bytes.Length);
            }

            var promoResponse = promoRequest.GetResponse() as HttpWebResponse;


            if (promoResponse != null && promoResponse.StatusCode == HttpStatusCode.OK)
                using (var s = promoResponse.GetResponseStream())
                {
                    // ReSharper disable once AssignNullToNotNullAttribute
                    using (var sr = new StreamReader(s, Encoding.GetEncoding(promoResponse.CharacterSet)))
                    {
                        result = JsonConvert.DeserializeObject<RedditAdJson>(sr.ReadToEnd());
                    }
                }

            return result;
        }

        /// <summary>
        /// Upload advertisement's image to Reddit's AWS S3 bucket
        /// </summary>
        /// <param name="adS3Data">Object containing all the info needed for the authentication to S3</param>
        /// <param name="imgPathText">Image's path</param>
        /// <param name="ad">Advertisement to be created</param>
        private void UploadImage(AdS3Json adS3Data, string imgPathText, Advertisement ad)
        {
            var imgPath = Path.Combine(imgPathText, ad.ThumbnailName);

            if (File.Exists(imgPath))
                using (var fs = new FileStream(imgPath, FileMode.Open, FileAccess.Read))
                {
                    var data = new byte[fs.Length];
                    fs.Read(data, 0, data.Length);

                    var postParameters = new Dictionary<string, object>
                                    {
                                        {"acl", "public-read"},
                                        {"key", adS3Data.Fields[1].Value},
                                        {"X-Amz-Credential", adS3Data.Fields[2].Value},
                                        {"X-Amz-Algorithm", adS3Data.Fields[3].Value},
                                        {"X-Amz-Date", adS3Data.Fields[4].Value},
                                        {"success_action_status", adS3Data.Fields[5].Value},
                                        {"content-type", adS3Data.Fields[6].Value},
                                        {"x-amz-storage-class", adS3Data.Fields[7].Value},
                                        {"x-amz-meta-ext", adS3Data.Fields[8].Value},
                                        {"policy", adS3Data.Fields[9].Value},
                                        {"X-Amz-Signature", adS3Data.Fields[10].Value},
                                        {"x-amz-security-token", adS3Data.Fields[11].Value},
                                        {
                                            "file",
                                            new FormUpload.FileParameter(data, ad.ThumbnailName,
                                                adS3Data.Fields[6].Value)
                                        }
                                    };
                    
                    FormUpload.MultipartFormDataPost("https://reddit-client-uploads.s3.amazonaws.com/",
                        "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36", postParameters);
                }
            else
                throw new FileNotFoundException("File does not exist on the specified directory!");
        }

        /// <summary>
        /// Queries Reddit to get the S3 data required for <see cref="UploadImage"/>
        /// </summary>
        /// <param name="ad">Advertisement to be created</param>
        /// <param name="adUh">Reddit's advertisement creation page modhash</param>
        /// <returns></returns>
        private AdS3Json GetAdS3Data(Advertisement ad, string adUh)
        {
            string adS3PostString = $"kind=thumbnail&link=&filepath={Uri.EscapeDataString(ad.ThumbnailName)}&uh={adUh}&ajax=true&raw_json=1";
            const string adS3Url = "https://www.reddit.com/api/ad_s3_params.json";
            var adS3Request = WebRequest.Create(adS3Url) as HttpWebRequest;
            AdS3Json adS3Data = new AdS3Json();

            if (adS3Request != null)
            {
                adS3Request.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
                adS3Request.Method = "POST";
                adS3Request.CookieContainer = Cookies;
                adS3Request.Accept = "application/json, text/javascript, */*; q=0.01";
                adS3Request.AutomaticDecompression = DecompressionMethods.GZip |
                                                     DecompressionMethods.Deflate;
                adS3Request.UserAgent =
                    "Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36";
                adS3Request.Referer = "https://www.reddit.com/";
                adS3Request.Timeout = 15000;

                var adS3CustomHeaders = adS3Request.Headers;

                adS3CustomHeaders.Add("accept-language", "en;q=0.4");
                adS3CustomHeaders.Add("origin", "https://www.reddit.com");
                adS3CustomHeaders.Add("x-requested-with", "XMLHttpRequest");

                var adS3Bytes = Encoding.ASCII.GetBytes(adS3PostString);

                adS3Request.ContentLength = adS3Bytes.Length;

                using (var os = adS3Request.GetRequestStream())
                {
                    os.Write(adS3Bytes, 0, adS3Bytes.Length);
                }

                var adS3Response = adS3Request.GetResponse() as HttpWebResponse;


                if (adS3Response != null && adS3Response.StatusCode == HttpStatusCode.OK)
                    using (var s = adS3Response.GetResponseStream())
                    {
                        if (s != null)
                            using (
                                // ReSharper disable once AssignNullToNotNullAttribute
                                var sr = new StreamReader(s, Encoding.GetEncoding(adS3Response.CharacterSet))
                            )
                            {
                                adS3Data = JsonConvert.DeserializeObject<AdS3Json>(sr.ReadToEnd());
                            }
                    }
            }

            return adS3Data;
        }

        /// <summary>
        /// Get a Reddit advertisement creation page modhash to create a new advertisement
        /// </summary>
        /// <returns>Reddit advertisement creation page modhash</returns>
        private string GetAdUh()
        {
            string adUh = string.Empty;

            var newPromoRequest = WebRequest.Create("https://www.reddit.com/promoted/new_promo/") as HttpWebRequest;

            if (newPromoRequest == null) throw new ArgumentNullException(nameof(newPromoRequest));

            newPromoRequest.CookieContainer = Cookies;
            newPromoRequest.Method = "GET";
            newPromoRequest.Accept =
                "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
            newPromoRequest.Timeout = 15000;

            var response = (HttpWebResponse)newPromoRequest.GetResponse();

            if (response.StatusCode == HttpStatusCode.OK)
                using (var s = response.GetResponseStream())
                {
                    if (s != null)
                        // ReSharper disable once AssignNullToNotNullAttribute
                        using (var sr = new StreamReader(s, Encoding.GetEncoding(response.CharacterSet)))
                        {
                            var doc = new HtmlDocument();

                            doc.Load(sr);

                            adUh =
                                doc.DocumentNode.SelectNodes("//form")[0].SelectNodes("//input")[0]
                                    .Attributes[2].Value;
                        }
                }

            return adUh;
        }
    }

    public class RedditAdJson
    {
        public List<List<object>> Jquery { get; set; }
        public bool Success { get; set; }
    }

    public class Field
    {
        public string Name { get; set; }
        public string Value { get; set; }
    }

    public class AdS3Json
    {
        public string Action { get; set; }
        public List<Field> Fields { get; set; }
    }
}