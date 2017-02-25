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

                if (error)
                {
                    MessageBox.Show(errorMsg, "Error!", MessageBoxButton.OK, MessageBoxImage.Error);
                }

            }
            else
                MessageBox.Show("Please select an image and Excel path before proceeding!", "ERROR", MessageBoxButton.OK, MessageBoxImage.Error);
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
    }
}
