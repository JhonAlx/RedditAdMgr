using MahApps.Metro.Controls;
using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;
using System.Windows;
using System.Threading.Tasks;

namespace RedditAdMgr
{
    /// <summary>
    /// Interaction logic for LoginWindow.xaml
    /// </summary>
    public partial class LoginWindow : MetroWindow
    {
        internal CookieContainer cookies { get; set; }

        public LoginWindow()
        {
            InitializeComponent();
            LoginProgressLabel.Visibility = Visibility.Hidden;
            LoginProgressBar.Visibility = Visibility.Hidden;
            UserNameTextBox.Text = Properties.Settings.Default.Username;
            PasswordTextBox.Password = Properties.Settings.Default.Password;

            if (!string.IsNullOrEmpty(Properties.Settings.Default.Username) && !string.IsNullOrEmpty(Properties.Settings.Default.Password))
                RememberPasswordCheckBox.IsChecked = true;
        }

        private void SubmitButton_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(UserNameTextBox.Text) || string.IsNullOrEmpty(PasswordTextBox.Password))
            {
                MessageBox.Show("Username or password fields are empty!");
            }
            else
            {
                var taskScheduler = TaskScheduler.FromCurrentSynchronizationContext();
                var ct = new CancellationToken();
                string c = string.Empty;
                string username = UserNameTextBox.Text, password = PasswordTextBox.Password;

                SubmitButton.IsEnabled = false;
                LoginProgressBar.Visibility = Visibility.Visible;
                LoginProgressBar.IsIndeterminate = true;
                LoginProgressLabel.Visibility = Visibility.Visible;
                LoginProgressLabel.Content = "Logging into Reddit...";

                Task.Factory.StartNew(() => Login(username, password, out c)).
                    ContinueWith(w =>
                   {
                       SubmitButton.IsEnabled = true;
                       LoginProgressBar.Visibility = Visibility.Hidden;
                       LoginProgressLabel.Visibility = Visibility.Hidden;

                       if (c.Contains("reddit_session"))
                       {
                           MessageBoxResult result = MessageBox.Show("Logged in!", "Success", MessageBoxButton.OK);

                           cookies = new CookieContainer();

                           cookies.SetCookies(new Uri("http://reddit.com"), c);

                           if (RememberPasswordCheckBox.IsChecked == true)
                           {
                               Properties.Settings.Default.Username = username;
                               Properties.Settings.Default.Password = password; //TODO: Encrypt this shit
                               Properties.Settings.Default.Save();
                           }

                           if (result == MessageBoxResult.OK)
                           {
                               MainForm main = new MainForm();
                               main.Cookies = cookies;
                               App.Current.MainWindow = main;
                               Close();
                               main.Show();
                           }
                       }
                       else
                           MessageBox.Show("Incorrect credentials, please try again!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                   },
                   ct,
                   TaskContinuationOptions.None, taskScheduler);
            }
        }

        private void Login(string username, string password, out string cookie)
        {
            string loginUrl = "https://www.reddit.com/api/login/";
            string loginParams = string.Format("op=login&user={0}&passwd={1}&api_type=json", username, password);

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
            cookie = response.Headers["set-cookie"];
        }
    }
}
