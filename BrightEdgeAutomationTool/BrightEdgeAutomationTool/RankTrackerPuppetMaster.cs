using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;

namespace BrightEdgeAutomationTool
{
    public static class RankTrackerPuppetMaster
    {
        private static User settings;
        public static void StartProcess(DirectoryInfo directory, User user)
        {
            if (!HWNDHelper.IsRankTrackerOpen())
            {
                System.Windows.MessageBox.Show("Please ensure Rank Tracker is open!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var rankTrackerHandle = HWNDHelper.GetRankTrackerWindow();
            if (rankTrackerHandle == IntPtr.Zero)
            {
                MessageBox.Show("Could not find Rank Tracker Window");
                return;
            }

            settings = user;

            FileInfo[] files = null;
            files = directory.GetFiles();

            foreach(FileInfo file in files)
            {
                SpreadsheetHelper.MatchedSheets.Clear();
                byte[] byteArray;
                try
                {
                    byteArray = File.ReadAllBytes(file.FullName);
                }
                catch (Exception e)
                {
                    //UpdateStatus($"{DateTime.Now} | Error reading file: {f.FullName}");
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
                            var resultSheetData = SpreadsheetHelper.GetResultSheetData(workbookPart);

                            var distinctResultSheetData = resultSheetData.Distinct(new DistinctKeywordResultComparer());

                            /*foreach (var res in distinctResultSheetData)
                            {
                                Console.WriteLine($"{res.Keyword} - {res.Volume}");
                            }*/

                            var keywordListStr = distinctResultSheetData.Select(x => x.Keyword).Aggregate((x, y) => x + "\n" + y);

                            /*HWNDHelper.BringWindowToFront(rankTrackerHandle);
                            HWNDHelper.SetRankTrackerSizeAndPosition(rankTrackerHandle);

                            RunRankTrackerProcess(keywordListStr);*/


                            // Process the Rank Tracker csv
                            var exportedFile = @"C:\Users\adria\OneDrive\Documents\Upwork\John Keyword\BrightEdgeAutomation\Rank Tracker Results Example.csv";

                            List<KeywordResultValue> RankTrackerKeywords = File.ReadAllLines(exportedFile)
                                .Skip(1).Select(v => KeywordResultValue.FromRankTrackerCsv(v))
                                .Where(v => v != null).ToList();
                            

                            distinctResultSheetData.ToList().ForEach(k => {
                                var item = RankTrackerKeywords.SingleOrDefault(l => l.Keyword.Equals(k.Keyword));
                                k.GoogleRank = item.GoogleRank;
                                k.RankingPage = item.RankingPage;
                            });

                            var finalResultSheetData = distinctResultSheetData.Select(d => d.RankingPage.Contains(mainSheetData.Marsha.ToLower()));

                            // Write final result sheet data to the results sheet
                        }
                    }
                }
            }
        }

        public static string RunRankTrackerProcess(string keywordListStr)
        {
            

            Clipboard.SetText(keywordListStr);

            // Click the add keyword button
            LeftMouseClick(382, 146);
            Thread.Sleep(1000);

            // Click in the keyword input area
            LeftMouseClick(335, 303);
            Thread.Sleep(1000);

            // Paste keywords
            PressPaste();

            // Click Next
            LeftMouseClick(686, 543);
            Thread.Sleep(2000);


            // Click Finish
            LeftMouseClick(770, 542);
            Thread.Sleep(2000);

            string pGreenColor = "#67B847"; // Progressbar green color //#61A032
            string pDarkColor = "#1F2530"; // Progressbar dark color

            var result = LoopUntil(() => {
                var c = HexConverter(GetColorAt(180, 596)); //184, 596
                if (!c.Equals(pGreenColor) && !c.Equals(pDarkColor))
                    return true;
                return false;
            }, TimeSpan.FromSeconds(60 * 1000 * 3)); // should be way longer in release

            Thread.Sleep(1000); // should be longer in release

            // Click Download button
            LeftMouseClick(1057, 145);
            Thread.Sleep(1000);

            // Save csv file
            //Keywords & rankings - test.csv
            var csvFileName = $"keywords_rankings_{DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond}.csv";
            SaveCsv(csvFileName);
            Thread.Sleep(1000);

            string exporterdFileName = "";
            
            // Get csv file
            var csvResult = LoopUntil(() => {
                var downloadedFiles = Directory.GetFiles(settings.RTExportPath)
                                        .Where(x => x.EndsWith(csvFileName));
                if(downloadedFiles.Count() > 0)
                {
                    exporterdFileName = downloadedFiles.FirstOrDefault();
                    return true;
                }


                return false;
            }, TimeSpan.FromMinutes(1));



            // Click in the keywords table and Delete keywords
            DeleteKeywords();

            return exporterdFileName;
        }





        public static bool LoopUntil(Func<bool> task, TimeSpan timeSpan)
        {
            bool success = false;
            int elapsed = 0;
            while ((!success) && (elapsed < timeSpan.TotalMilliseconds))
            {
                Thread.Sleep(100);
                elapsed += 100;
                success = task();
            }
            return success;
        }




        private static String HexConverter(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        private static String RGBConverter(System.Drawing.Color c)
        {
            return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetDesktopWindow();
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindowDC(IntPtr window);
        [DllImport("gdi32.dll", SetLastError = true)]
        public static extern uint GetPixel(IntPtr dc, int x, int y);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int ReleaseDC(IntPtr window, IntPtr dc);

        public static Color GetColorAt(int x, int y)
        {
            IntPtr desk = GetDesktopWindow();
            IntPtr dc = GetWindowDC(desk);
            int a = (int)GetPixel(dc, x, y);
            ReleaseDC(desk, dc);
            return Color.FromArgb(255, (a >> 0) & 0xff, (a >> 8) & 0xff, (a >> 16) & 0xff);
        }







        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);

        public const int KEYEVENTF_KEYDOWN = 0x0000; // New definition
        public const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        public const int VK_LCONTROL = 0xA2; //Left Control key code
        public const int VK_MENU = 0x12; //ALT key code
        public const int VK_DELETE = 0x2E; //DEL key code
        public const int VK_RETURN = 0x0D; //ENTER key code

        public const int A = 0x41; //A key code
        public const int C = 0x43; //C key code
        public const int N = 0x4E; //N key code
        public const int V = 0x56; //V key code
        public const int Y = 0x59; //Y key code

        public static void PressPaste()
        {
            // Hold Control down and press V
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(V, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(V, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void SaveCsv(string filename)
        {
            AltN(); Thread.Sleep(500);
            CtrlA(); Thread.Sleep(500);

            Clipboard.SetText(filename);
            PressPaste(); Thread.Sleep(500);
            PressEnter(); Thread.Sleep(500);
            // Do not open folder
            AltN(); Thread.Sleep(500);
        }

        public static void DeleteKeywords()
        {
            LeftMouseClick(600, 222); Thread.Sleep(500);
            CtrlA(); Thread.Sleep(500);
            PressDelete(); Thread.Sleep(500);
            AltY(); Thread.Sleep(500);
        }

        public static void AltN()
        {
            // Hold ALT down and press N
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(N, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(N, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void AltY()
        {
            // Hold ALT down and press N
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(Y, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(Y, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void CtrlA()
        {
            // Hold Ctrl down and press A
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void PressKeys()
        {
            // Hold Control down and press A
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(A, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);

            // Hold Control down and press C
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(C, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(C, 0, KEYEVENTF_KEYUP, 0);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }


        public static void PressEnter()
        {
            keybd_event(VK_RETURN, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void PressDelete()
        {
            keybd_event(VK_DELETE, 0, KEYEVENTF_KEYDOWN, 0);
            keybd_event(VK_DELETE, 0, KEYEVENTF_KEYUP, 0);
        }



        //This is a replacement for Cursor.Position in WinForms
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        static extern bool SetCursorPos(int x, int y);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_LEFTDOWN = 0x02;
        public const int MOUSEEVENTF_LEFTUP = 0x04;

        //This simulates a left mouse click
        public static void LeftMouseClick(int xpos, int ypos)
        {
            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);
        }



    }

    class DistinctKeywordResultComparer : IEqualityComparer<KeywordResultValue>
    {

        public bool Equals(KeywordResultValue x, KeywordResultValue y)
        {
            return x.Keyword == y.Keyword && x.Volume == y.Volume;
        }

        public int GetHashCode(KeywordResultValue obj)
        {
            return obj.Keyword.GetHashCode() ^
                obj.Volume.GetHashCode();
        }
    }
}
