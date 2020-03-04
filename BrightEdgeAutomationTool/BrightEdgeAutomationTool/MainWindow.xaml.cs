using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace BrightEdgeAutomationTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static IWebDriver Driver { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }



        private string GetPathToChrome()
        {
            const string suffix = @"Google\Chrome\Application\chrome.exe";
            var prefixes = new List<string> { Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) };
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            var programFilesx86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (programFilesx86 != programFiles)
            {
                prefixes.Add(programFiles);
            }
            else
            {
                var programFilesDirFromReg = Microsoft.Win32.Registry.GetValue(@"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion", "ProgramW6432Dir", null) as string;
                if (programFilesDirFromReg != null) prefixes.Add(programFilesDirFromReg);
            }

            prefixes.Add(programFilesx86);
            var path = prefixes.Distinct().Select(prefix => System.IO.Path.Combine(prefix, suffix)).FirstOrDefault(File.Exists);

            return path;
        }

        private void launchChrome_Click(object sender, RoutedEventArgs e)
        {
            launchChrome.IsEnabled = false;

            Thread chromeThread = new Thread(() =>
            {
                string pathToChrome = "";
                pathToChrome = GetPathToChrome();

                if (pathToChrome == "")
                {
                    System.Windows.MessageBox.Show("Unable to locate Chrome!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }


                Process proc = new Process();
                proc.StartInfo.FileName = pathToChrome;

                proc.StartInfo.Arguments = "https://www.brightedge.com/secure/login/ --new-window --remote-debugging-port=9222 --user-data-dir=C:\\Temp";
                proc.Start();

                var chromeDriverService = ChromeDriverService.CreateDefaultService();
                chromeDriverService.HideCommandPromptWindow = true;
                //return new ChromeDriver(chromeDriverService, new ChromeOptions());

                ChromeOptions options = new ChromeOptions();
                options.DebuggerAddress = "127.0.0.1:9222";


                options.AddArgument("--start-maximized");
                options.AddArguments("--disable-gpu");
                options.AddArguments("--disable-extensions");

                Driver = new ChromeDriver(chromeDriverService, options);
                //Driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);
                PuppetMaster.Driver = Driver;
                PuppetMaster.Window = this;

                PuppetMaster.LoginUser();

                this.Dispatcher.Invoke(() =>
                {
                    start.IsEnabled = true;
                });

                while (proc.HasExited == false)
                {
                    if ((DateTime.Now.Second % 5) == 0)
                    { // Show a tick every five seconds.
                        //Console.Write(".");
                    }
                    System.Threading.Thread.Sleep(1000);
                }

                // After process exits
                try
                {
                    this.Dispatcher.Invoke(() =>
                    {
                        launchChrome.IsEnabled = true;
                        start.IsEnabled = false;
                    });
                }
                catch (Exception ex)
                { }

            });

            chromeThread.Start();


        }

        private void start_Click(object sender, RoutedEventArgs e)
        {

        }

        private void stopProcess_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
