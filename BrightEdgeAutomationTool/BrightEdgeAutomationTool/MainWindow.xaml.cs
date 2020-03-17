using DocumentFormat.OpenXml.Packaging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media.Animation;
using Path = System.IO.Path;

namespace BrightEdgeAutomationTool
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static IWebDriver Driver { get; set; }
        private string DownloadsFolder;
        private List<string> FilesToDelete = new List<string>();
        //private Database database;
        private DbCreator database;
        private User user;
        public bool StopProcess { get; set; } = false;

        public MainWindow()
        {
            InitializeComponent();



            SHGetKnownFolderPath(KnownFolder.Downloads, 0, IntPtr.Zero, out DownloadsFolder);

            database = new DbCreator();
            user = database.GetUser();
            AddUserFields();

            CheckIfRankTrackerIsOpen();
            /*Thread t = new Thread(() =>
            {
                CheckIfRankTrackerIsOpen();
            });

            t.Start();*/

            // Disable start button initially 
            SetButtonState();


        }

        private void SetButtonState()
        {
            if (user.RunRankTracker == true && user.RunBrightEdge == true)
            {
                start.IsEnabled = false;
                launchChrome.IsEnabled = true;
            }
            else if (user.RunRankTracker == true && user.RunBrightEdge == false)
            {
                start.IsEnabled = true;
                launchChrome.IsEnabled = false;
            }
            else if (user.RunRankTracker == false && user.RunBrightEdge == true)
            {
                start.IsEnabled = false;
                launchChrome.IsEnabled = true;
            }
            else
            {
                start.IsEnabled = false;
                launchChrome.IsEnabled = false;
            }
        }

        private void CheckIfRankTrackerIsOpen()
        {
            if (user.RunRankTracker)
            {
                if (!HWNDHelper.IsRankTrackerOpen())
                {
                    System.Windows.MessageBox.Show("Please ensure Rank Tracker is open!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void AddUserFields()
        {
            email.Text = user.Email;
            password.Password = user.Password;
            bright_edge.IsChecked = user.RunBrightEdge;
            rank_tracker.IsChecked = user.RunRankTracker;
            rt_export_path.Text = user.RTExportPath;
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

                PuppetMaster.LoginUser(user);

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
            /*var keywords =
                @"downtown puchong city center meeting facilities
meeting facility downtown puchong city center
meeting facilities downtown puchong city center
meeting facility in downtown puchong city center
meeting facilities in downtown puchong city center
downtown puchong financial corporate centre meeting hotel
downtown puchong financial corporate centre meeting hotels
meeting hotel downtown puchong financial corporate centre
meeting hotels downtown puchong financial corporate centre
meeting hotel in downtown puchong financial corporate centre";

            var res = PuppetMaster.GetFirstRowLinkWithParam(keywords, "Malaysia");

            return;*/

            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {

                    Thread puppetThread = new Thread(() =>
                    {
                        try
                        {
                            PowerHelper.ForceSystemAwake();
                        }
                        catch (Exception ex) { }

                        //PuppetMaster.baseURL = Driver.Url;
                        StartPuppetProcess(fbd.SelectedPath);

                        try
                        {
                            PowerHelper.ResetSystemDefault();
                        }
                        catch (Exception ex) { }

                    });
                    puppetThread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
                    puppetThread.Start();
                }
            }
        }

        private void StartPuppetProcess(string selectedPath)
        {

            this.Dispatcher.Invoke(() =>
            {
                start.IsEnabled = false;
                StopProcess = false;
                stopProcess.IsEnabled = true;

                status.Text = "";
                SpinnerText.Visibility = Visibility.Visible;
            });

            DirectoryInfo dirInfo = new DirectoryInfo(selectedPath);
            FileInfo[] files = null;
            files = dirInfo.GetFiles();
            int batchesProcessed = 0;
            DirectoryInfo dirForRT = null;

            // Start Bright Edge Processing
            if (user.RunBrightEdge)
            {
                foreach (FileInfo f in files)
                {

                    string fileToProcess = f.FullName;

                    if (f.Name.StartsWith("~$"))
                        continue;

                    UpdateStatus($"{DateTime.Now} | Processing file {f.Name}");
                    SpreadsheetHelper.MatchedSheets.Clear();
                    byte[] byteArray;
                    try
                    {
                        byteArray = File.ReadAllBytes(fileToProcess);
                    }
                    catch (Exception e)
                    {
                        UpdateStatus($"{DateTime.Now} | Error reading file: {f.Name}");
                        continue;
                    }

                    using (MemoryStream stream = new MemoryStream())
                    {
                        stream.Write(byteArray, 0, (int)byteArray.Length);

                        // Open the document for editing
                        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true))
                        {
                            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                            SpreadsheetHelper.workbookPart = workbookPart;

                            var mainSheetPart = SpreadsheetHelper.GetWorksheetPart(workbookPart, "REPLACE");

                            if (mainSheetPart != null)
                            {
                                var mainSheetData = SpreadsheetHelper.GetMainSheetData(mainSheetPart);
                                List<KeywordResultValue> keywordStats = new List<KeywordResultValue>();

                                // Process all Pages
                                foreach (var pageItem in mainSheetData.Pages)
                                {
                                    // Only for testing and should be commented out **********************
                                    //if (pageItem != "Dining")
                                    //    continue;

                                    UpdateStatus($"{DateTime.Now} | Processing {pageItem} keywords in file: {f.Name}");

                                    //Console.WriteLine(item);
                                    var keywordListData = SpreadsheetHelper.GetKeywordsFromSheet(pageItem);
                                    var keywordList = keywordListData.Item1;
                                    var keywordCount = keywordListData.Item2;

                                    UpdateStatus($"{DateTime.Now} | {pageItem}: keyword count: {keywordCount}");

                                    PuppetMaster.RetryUntilSuccessOrTimeout(() =>
                                    {


                                        List<KeywordResultValue> keywordPageStats = new List<KeywordResultValue>();

                                        // Process a 1000 keywords at a time from each page
                                        foreach (var keywordListItem in keywordList)
                                        {

                                            if (StopProcess)
                                                return true;

                                            DateTime processStartTime = DateTime.Now;

                                            var result = PuppetMaster.RunProcess(keywordListItem, mainSheetData.Country);

                                            if (result == false)
                                                return false;


                                            // Process downloaded file
                                            IEnumerable<string> downloadedFiles = new List<string>();

                                            var diffInSeconds = (processStartTime - DateTime.Now).TotalSeconds;

                                            while (diffInSeconds <= 60 && downloadedFiles.Count() == 0)
                                            {
                                                downloadedFiles = Directory.GetFiles(DownloadsFolder)
                                                    .Where(x => new FileInfo(x).CreationTime > processStartTime && x.EndsWith(".csv"));

                                                Thread.Sleep(3000);
                                            }

                                            if (downloadedFiles.Count() > 0)
                                            {
                                                try
                                                {
                                                    // Process the file
                                                    var downloadedFile = downloadedFiles.ElementAt(0);
                                                    Console.WriteLine(downloadedFile);
                                                    FilesToDelete.Add(downloadedFile);

                                                    List<KeywordResultValue> keywordStats1000 = File.ReadAllLines(downloadedFile)
                                                        .Skip(1).Select(v => KeywordResultValue.FromCsv(v))
                                                        .Where(v => v != null).ToList();

                                                    keywordPageStats.AddRange(keywordStats1000);
                                                }
                                                catch (Exception ex)
                                                {
                                                    return false;
                                                }
                                            }
                                            else
                                            {
                                                return false;
                                            }


                                        }

                                        //if (keywordPageStats.Count() == 0)
                                        //    return false;

                                        keywordPageStats = keywordPageStats.OrderByDescending(k => k.Volume).ToList();

                                        UpdateStatus($"{DateTime.Now} | {pageItem}: volume stats found in csv: {keywordPageStats.Count}");

                                        keywordStats.AddRange(keywordPageStats);

                                        DeleteDownloadedFiles();

                                        return true;
                                    }, TimeSpan.FromMinutes(8));

                                    //break;
                                }

                                PuppetMaster.RemoveLocation(mainSheetData.Country);


                                SpreadsheetHelper.CreateResultSheet(keywordStats);
                                batchesProcessed++;
                            }


                        }

                        var saveResult = SaveAs(fileToProcess, stream);

                        dirForRT = saveResult.Item2;
                    }

                    if (StopProcess)
                        break;
                }

                PuppetMaster.DeleteAllQueries((int)Math.Ceiling((decimal)batchesProcessed / 10));
            }


            // Start Rank Tracker Processing
            if (user.RunRankTracker)
            {
                RankTrackerPuppetMaster.MainWindow = this;


                if (dirForRT == null)
                {

                    dirForRT = new DirectoryInfo(selectedPath);

                    /*using (var fbd = new FolderBrowserDialog())
                    {
                        DialogResult result = fbd.ShowDialog();

                        if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                        {
                            dirForRT = new DirectoryInfo(fbd.SelectedPath);

                            // Start the process
                            RankTrackerPuppetMaster.StartProcess(dirForRT);

                            dirForRT = null;
                        }
                    }*/
                }

                // Start the process
                RankTrackerPuppetMaster.StartProcess(dirForRT, user);

                dirForRT = null;
            }






            this.Dispatcher.Invoke(() =>
            {
                start.IsEnabled = true;
                SpinnerText.Visibility = Visibility.Collapsed;
                stopProcess.IsEnabled = false;
            });
            UpdateStatus($"{DateTime.Now} | Process complete!");
            System.Windows.MessageBox.Show("Process complete!", "Complete", MessageBoxButton.OK, MessageBoxImage.Information);


        }

        private void stopProcess_Click(object sender, RoutedEventArgs e)
        {
            UpdateStatus($"{DateTime.Now} | Cancelling the process...");
            StopProcess = true;
            stopProcess.IsEnabled = false;
        }

        public void UpdateStatus(string log)
        {
            this.Dispatcher.Invoke(() =>
            {
                status.Text = status.Text + Environment.NewLine + log;
            });
        }

        public void DeleteDownloadedFiles()
        {
            foreach (var value in FilesToDelete)
            {
                //Delete files when done
                if (File.Exists(value))
                {
                    File.Delete(value);
                }
            }
        }

        private (bool, DirectoryInfo) SaveAs(string path, MemoryStream stream)
        {
            var directory = Path.GetDirectoryName(path);
            var fileName = Path.GetFileName(path);


            var newPath = Path.Combine(directory, "results");
            Directory.CreateDirectory(newPath);

            var newFilePath = Path.Combine(newPath, fileName);

            File.WriteAllBytes(newFilePath, stream.ToArray());



            return (true, new DirectoryInfo(newPath));
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                Driver.Close();
                Driver.Dispose();
                database.Close();

                Process[] chromeDriverProcesses = Process.GetProcessesByName("chromedriver");

                foreach (var chromeDriverProcess in chromeDriverProcesses)
                {
                    chromeDriverProcess.Kill();
                }
            }
            catch (Exception ex)
            {

            }

        }

        public static class KnownFolder
        {
            public static readonly Guid Downloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out string pszPath);

        public bool IsRightMenuVisible { get; set; } = false;
        private void BtnRightMenuShow_Click(object sender, RoutedEventArgs e)
        {
            if (IsRightMenuVisible)
            {
                ShowHideMenu("sbHideRightMenu", pnlRightMenu);
                overlay.Visibility = Visibility.Collapsed;
                IsRightMenuVisible = false;
            }
            else
            {
                user = database.GetUser();
                AddUserFields();

                ShowHideMenu("sbShowRightMenu", pnlRightMenu);
                overlay.Visibility = Visibility.Visible;
                IsRightMenuVisible = true;
            }
        }


        private void ShowHideMenu(string storyboard, StackPanel pnl)
        {
            Storyboard sb = Resources[storyboard] as Storyboard;
            sb.Begin(pnl);
        }



        private void SaveSettings_Click(object sender, RoutedEventArgs e)
        {
            var emailValue = email.Text;
            var passwordValue = password.Password.ToString();
            var BEValue = (bool)bright_edge.IsChecked;
            var RTValue = (bool)rank_tracker.IsChecked;
            var RTExportPathValue = rt_export_path.Text;

            if (emailValue == "" || passwordValue == "")
            {
                System.Windows.MessageBox.Show("All fields are required", "Warning",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var user = database.UpdateUser(emailValue, passwordValue, BEValue, RTValue, RTExportPathValue);

            if (user != null)
            {
                success.Visibility = Visibility.Visible;
                var timer = new System.Windows.Forms.Timer();
                timer.Interval = 2000; // here time in milliseconds
                timer.Tick += (object s, System.EventArgs ev) =>
                {
                    success.Visibility = Visibility.Collapsed;
                };
                timer.Start();
                this.user = user;

                SetButtonState();

                return;
            }

            System.Windows.MessageBox.Show("Error while saving", "Error",
                   MessageBoxButton.OK, MessageBoxImage.Error);
        }

        private void Overlay_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            if (IsRightMenuVisible)
            {
                ShowHideMenu("sbHideRightMenu", pnlRightMenu);
                overlay.Visibility = Visibility.Collapsed;
                IsRightMenuVisible = false;
            }
            else
            {
                ShowHideMenu("sbShowRightMenu", pnlRightMenu);
                overlay.Visibility = Visibility.Visible;
                IsRightMenuVisible = true;
            }
        }

        private void setRTExportPath_Click(object sender, RoutedEventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                DialogResult result = fbd.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK && !string.IsNullOrWhiteSpace(fbd.SelectedPath))
                {
                    rt_export_path.Text = fbd.SelectedPath;
                }
            }
        }
    }
}
