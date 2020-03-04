﻿using DocumentFormat.OpenXml.Packaging;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
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
        private string DownloadsFolder;
        private List<string> FilesToDelete = new List<string>();

        public MainWindow()
        {
            InitializeComponent();


            SHGetKnownFolderPath(KnownFolder.Downloads, 0, IntPtr.Zero, out DownloadsFolder);
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
                //StopProcess = false;
                stopProcess.IsEnabled = true;

                status.Text = "";
                SpinnerText.Visibility = Visibility.Visible;
            });

            DirectoryInfo dirInfo = new DirectoryInfo(selectedPath);
            FileInfo[] files = null;
            files = dirInfo.GetFiles();




            foreach (FileInfo f in files)
            {

                string fileToProcess = f.FullName;


                byte[] byteArray = File.ReadAllBytes(fileToProcess);
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

                            //Console.WriteLine(mainSheetData.Marsha);
                            //Console.WriteLine(mainSheetData.Country);

                            // Process all Pages
                            foreach (var item in mainSheetData.Pages)
                            {

                                List<KeywordResultValue> keywordPageStats = new List<KeywordResultValue>();

                                //Console.WriteLine(item);
                                var keywordList = SpreadsheetHelper.GetKeywordsFromSheet(item);
                                // Process a 1000 keywords at a time from eage page
                                foreach (var keywordListItem in keywordList)
                                {
                                    //Console.WriteLine(keywordList.Count);
                                    //Console.WriteLine(keywordListItem);

                                    DateTime processStartTime = DateTime.Now;

                                    PuppetMaster.RunProcess(keywordListItem, mainSheetData.Country);

                                    // Process downloaded file
                                    IEnumerable<string> downloadedFiles = new List<string>();

                                    var diffInSeconds = (processStartTime - DateTime.Now).TotalSeconds;

                                    while (diffInSeconds <= 60 && downloadedFiles.Count() == 0)
                                    {
                                        downloadedFiles = Directory.GetFiles(DownloadsFolder)
                                            .Where(x => new FileInfo(x).CreationTime > processStartTime);

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

                                            //keywordStatsByVolume = keywordStatsByVolume.OrderByDescending(k => k.Volume).ToList();

                                            keywordPageStats.AddRange(keywordStats1000);


                                            //keywordStats.AddRange(keywordStatsByVolume);
                                            //keywordStats = keywordStats.Concat(keywordStatsByVolume);
                                        }
                                        catch (Exception ex)
                                        {
                                        }
                                        finally
                                        {
                                            
                                        }

                                    }
                                }

                                keywordPageStats = keywordPageStats.OrderByDescending(k => k.Volume).ToList();
                                keywordStats.AddRange(keywordPageStats);

                                DeleteDownloadedFiles();

                                break;
                            }



                        }
                    }
                }
            }

            

        }

        private void stopProcess_Click(object sender, RoutedEventArgs e)
        {

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

        public static class KnownFolder
        {
            public static readonly Guid Downloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out string pszPath);
    }
}
