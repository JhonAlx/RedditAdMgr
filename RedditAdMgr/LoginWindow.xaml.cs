using MahApps.Metro.Controls;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Windows;

namespace RedditAdMgr
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : MetroWindow
    {
        public CookieContainer cookies { get; set; }

        public LoginWindow()
        {
            InitializeComponent();
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
        }

        private void Login()
        {
            string loginUrl = "https://www.reddit.com/api/login/";
            string loginParams = string.Format("op=login&user={0}&passwd={1}&api_type=json", UserNameTextBox.Text, PasswordTextBox.Password);

            //statusLbl.Visibility = Visibility.Visible;
            //progressBar.Visibility = Visibility.Visible;

            //statusLbl.Content = "Logging in into Reddit...";

            HttpWebRequest loginRequest = WebRequest.Create(loginUrl) as HttpWebRequest;
            loginRequest.ContentType = "application/x-www-form-urlencoded";
            loginRequest.Method = "POST";

            byte[] bytes = Encoding.ASCII.GetBytes(loginParams);

            loginRequest.ContentLength = bytes.Length;

            using (Stream os = loginRequest.GetRequestStream())
            {
                os.Write(bytes, 0, bytes.Length);
            }

            HttpWebResponse response = loginRequest.GetResponse() as HttpWebResponse;
            string cookie = response.Headers["set-cookie"];

            //statusLbl.Content = "Logged in!";

            cookies = new CookieContainer();

            cookies.SetCookies(new Uri("http://reddit.com"), cookie);
        }
    }
}
