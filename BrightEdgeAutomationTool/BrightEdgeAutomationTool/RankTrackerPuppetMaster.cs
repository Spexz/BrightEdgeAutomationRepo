using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;
using System.Windows.Forms;

namespace BrightEdgeAutomationTool
{
    public static class RankTrackerPuppetMaster
    {
        private static User settings;
        public static MainWindow MainWindow;
        private static RTSettings rtSettings;
        public static void StartProcess(DirectoryInfo directory, User user)
        {
            rtSettings = MainWindow.database.GetRTSettings();

            if (!HWNDHelper.IsRankTrackerOpen())
            {
                System.Windows.MessageBox.Show("Please ensure Rank Tracker is open!", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var rankTrackerHandle = HWNDHelper.GetRankTrackerWindow();
            if (rankTrackerHandle == IntPtr.Zero)
            {
                System.Windows.MessageBox.Show("Could not find Rank Tracker Window");
                return;
            }

            if (rtSettings == null)
            {
                System.Windows.MessageBox.Show("Please ensure that the Rank Tracker Settings are done", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            MainWindow.UpdateStatus($"{DateTime.Now} | Starting Rank Tracker process");

            settings = user;

            FileInfo[] files = null;
            files = directory.GetFiles();

            foreach (FileInfo file in files)
            {

                MainWindow.UpdateStatus($"{DateTime.Now} | Processing file {file.Name}");

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
                            string rankTrackerCsvFile = null;

                            var result = LoopUntil(() =>
                            {

                                HWNDHelper.BringWindowToFront(rankTrackerHandle);
                                Thread.Sleep(1000);
                                HWNDHelper.SetRankTrackerSizeAndPosition(rankTrackerHandle, (msg) =>
                                {
                                    MainWindow.UpdateStatus(msg);
                                    return true;
                                }, rtSettings);



                                rankTrackerCsvFile = RunRankTrackerProcess(keywordListStr);
                                //return true; // to be removed

                                if (String.IsNullOrEmpty(rankTrackerCsvFile))
                                {
                                    PressEscape();
                                    Thread.Sleep(500);
                                    PressEscape();
                                    DeleteKeywords();

                                    return false;
                                }

                                return true;
                            }, TimeSpan.FromMinutes(60));


                            //return; // to be removed



                            // Process the Rank Tracker csv // to be removed
                            //rankTrackerCsvFile = @"C:\Users\adria\OneDrive\Desktop\Test RT\Rank Tracker Results Example.csv";

                            List<KeywordResultValue> RankTrackerKeywords = File.ReadAllLines(rankTrackerCsvFile)
                                .Skip(1).Select(v => KeywordResultValue.FromRankTrackerCsv(v))
                                .Where(v => v != null).ToList();

                            //Console.WriteLine(rankTrackerCsvFile);


                            distinctResultSheetData.ToList().ForEach(k =>
                            {
                                //var item = RankTrackerKeywords.SingleOrDefault(l => l.Keyword.Equals(k.Keyword));
                                var item = RankTrackerKeywords.SingleOrDefault(l =>
                                String.Compare(l.Keyword, k.Keyword, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase | CompareOptions.IgnoreSymbols) == 0);
                                //String.Compare(s1, s2, CultureInfo.CurrentCulture, CompareOptions.IgnoreCase | CompareOptions.IgnoreSymbols) == 0
                                if (item != null)
                                {
                                    k.GoogleRank = item.GoogleRank;
                                    k.RankingPage = item.RankingPage;
                                }

                            });

                            var finalResultSheetData = distinctResultSheetData.Where(d => d.RankingPage == null ? false :
                                d.RankingPage.Contains(mainSheetData.Marsha.ToLower())).ToList();

                            MainWindow.UpdateStatus($"{DateTime.Now} | Rank Tracker: Pulled {finalResultSheetData.Count()} Keywords from {file.Name}");

                            // Write final result sheet data to the results sheet

                            /*foreach (var res in finalResultSheetData)
                            {
                                Console.WriteLine($"{res.Keyword} - {res.Volume} - {res.GoogleRank} - {res.RankingPage}");
                            }*/

                            SpreadsheetHelper.RTUpdateResultSheet(workbookPart, finalResultSheetData);

                            spreadsheetDocument.WorkbookPart.Workbook.Save();

                            /*if (File.Exists(rankTrackerCsvFile))
                            {
                                File.Delete(rankTrackerCsvFile);
                            }*/
                        }
                    }

                    // Save
                    File.WriteAllBytes(file.FullName, stream.ToArray());
                }
            }
        }

        public static string RunRankTrackerProcess(string keywordListStr)
        {
            //return ""; // to be removed
            
            var rtHwnd = HWNDHelper.GetRankTrackerWindow();

            System.Windows.Clipboard.SetText(keywordListStr);
            Thread.Sleep(2000);

            HWNDHelper.WaitForWindowToRespond(rtHwnd);
            WaitForCursor();

            // Click the add keyword button
            //LeftMouseClick(565, 137);
            LeftMouseClick(rtSettings.AddKeywordPos.X, rtSettings.AddKeywordPos.Y);
            Thread.Sleep(2000);
            WaitForCursor();
            //HWNDHelper.FindAndBringFwd("Add New Keywords");

            // Click in the keyword input area
            //LeftMouseClick(590, 337);
            //LeftMouseClick(rtSettings, 337);
            Thread.Sleep(2000);
            WaitForCursor();

            // Paste keywords
            PressPaste();
            Thread.Sleep(2000);
            WaitForCursor();

            // Click Next
            //LeftMouseClick(811, 575);
            LeftMouseClick(rtSettings.NextPos.X, rtSettings.NextPos.Y);
            Thread.Sleep(2000);
            WaitForCursor();


            // Click Finish
            //LeftMouseClick(919, 575);
            LeftMouseClick(rtSettings.FinishPos.X, rtSettings.FinishPos.Y);
            Thread.Sleep(2000);
            WaitForCursor();

            string pGreenColor = "#67B847"; // Progressbar green color
            string pGreenColor2 = "#61A032"; // Progressbar green color
            string pDarkColor = "#1F2530"; // Progressbar dark color

            List<string> colors = new List<string>()
            {
                pGreenColor, pGreenColor2, pDarkColor
            };

            var result = LoopUntil(() =>
            {
                //var c = HexConverter(GetColorAt(134, 656));
                /*var c = HexConverter(GetColorAt(rtSettings.ProgressbarPos.X, rtSettings.ProgressbarPos.Y));
                if (!c.Equals(pGreenColor) && !c.Equals(pDarkColor) && !c.Equals(pGreenColor2))
                    return true;*/

                if (!IsColorAlongYAxis(rtSettings.ProgressbarPos, colors, 5))
                    return true;

                return false;
            }, TimeSpan.FromSeconds(60 * 1000 * 5)); // should be way longer in release

            Thread.Sleep(2000); // should be longer in release

            // Click Download button
            //LeftMouseClick(1310, 137);
            LeftMouseClick(rtSettings.DownloadPos.X, rtSettings.DownloadPos.Y);
            Thread.Sleep(2000);
            WaitForCursor();

#if DEBUG
            //For trial version only
            //LeftMouseClick(981, 128);
            PressEscape();
            Thread.Sleep(1000);
            WaitForCursor();

            PressEnter();
            Thread.Sleep(1000);
            WaitForCursor();
#endif


            // Save csv file
            //Keywords & rankings - test.csv
            var csvFileName = $"keywords_rankings_{DateTime.Now.Ticks / TimeSpan.TicksPerMillisecond}.csv";
            SaveCsv(csvFileName);
            Thread.Sleep(1000);
            WaitForCursor();

            string exportedFileName = "";

            // Get csv file
            var csvResult = LoopUntil(() =>
            {
                var downloadedFiles = Directory.GetFiles(settings.RTExportPath)
                                        .Where(x => x.EndsWith(csvFileName));
                if (downloadedFiles.Count() > 0)
                {
                    exportedFileName = downloadedFiles.FirstOrDefault();
                    return true;
                }


                return false;
            }, TimeSpan.FromMinutes(2));



            // Click in the keywords table and Delete keywords
            DeleteKeywords();

            return exportedFileName;
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


        private static bool IsColorAlongYAxis(System.Drawing.Point point, List<string> colors, int yOffset)
        {
            var yStart = point.Y - yOffset;
            var yEnd = point.Y + yOffset;

            for(var i = yStart; i <= yEnd; i++)
            {
                var c = HexConverter(GetColorAt(point.X, i));
                if (colors.Contains(c))
                {
                    return true;
                }
            }

            return false;
        }



        public static String HexConverter(System.Drawing.Color c)
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
        public const int VK_ESCAPE = 0x1B; //Escape key code

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
            //HWNDHelper.FindAndBringFwd("Save");
            AltN(); Thread.Sleep(1000);
            WaitForCursor();
            CtrlA(); Thread.Sleep(1000);
            WaitForCursor();

            System.Windows.Clipboard.SetText(filename);
            PressPaste(); Thread.Sleep(1000);
            WaitForCursor();
            PressEnter(); Thread.Sleep(3000);
            WaitForCursor();
            // Do not open folder
            AltN(); Thread.Sleep(1000);
        }

        public static void DeleteKeywords()
        {
            // Click in the keyword table
            //LeftMouseClick(718, 213); Thread.Sleep(1000);
            LeftMouseClick(rtSettings.KeywordInputPos.X, rtSettings.KeywordInputPos.Y);
            //WaitForCursor();

            CtrlA(); Thread.Sleep(2000);
            //WaitForCursor();

            PressDelete(); Thread.Sleep(3000);
            //WaitForCursor();
            //HWNDHelper.FindAndBringFwd("Removal confirmation");

            AltY(); Thread.Sleep(1000);
        }

        public static void AltN()
        {
            // Hold ALT down and press N
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(100);
            keybd_event(N, 0, KEYEVENTF_KEYDOWN, 0);

            Thread.Sleep(300);

            keybd_event(N, 0, KEYEVENTF_KEYUP, 0);
            Thread.Sleep(100);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void AltY()
        {
            // Hold ALT down and press Y
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(100);
            keybd_event(Y, 0, KEYEVENTF_KEYDOWN, 0);

            Thread.Sleep(300);

            keybd_event(Y, 0, KEYEVENTF_KEYUP, 0);
            Thread.Sleep(100);
            keybd_event(VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void CtrlA()
        {
            // Hold Ctrl down and press A
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(100);
            keybd_event(A, 0, KEYEVENTF_KEYDOWN, 0);

            Thread.Sleep(300);

            keybd_event(A, 0, KEYEVENTF_KEYUP, 0);
            Thread.Sleep(100);
            keybd_event(VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
        }


        public static void PressEnter()
        {
            keybd_event(VK_RETURN, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(100);
            keybd_event(VK_RETURN, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void PressDelete()
        {
            keybd_event(VK_DELETE, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(100);
            keybd_event(VK_DELETE, 0, KEYEVENTF_KEYUP, 0);
        }

        public static void PressEscape()
        {
            keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYDOWN, 0);
            Thread.Sleep(100);
            keybd_event(VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0);
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
            var result = HWNDHelper.BlockInput(true);

            //Console.WriteLine($"Block input: {result}");

            SetCursorPos(xpos, ypos);
            mouse_event(MOUSEEVENTF_LEFTDOWN, xpos, ypos, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP, xpos, ypos, 0, 0);

            HWNDHelper.BlockInput(false);
        }


        private static void WaitForCursor()
        {
            LoopUntil(() =>
            {
                if (IsWaitCursor() == false)
                    return true;

                return false;
            }, TimeSpan.FromMinutes(5));
        }

        private static bool IsWaitCursor()
        {
            var h = Cursors.WaitCursor.Handle;

            CURSORINFO pci;
            pci.cbSize = Marshal.SizeOf(typeof(CURSORINFO));
            GetCursorInfo(out pci);
            //Console.WriteLine(pci.hCursor.ToString());
            //Console.WriteLine(pci.hCursor == h);
            return pci.hCursor == h;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct POINT
        {
            public Int32 x;
            public Int32 y;
        }

        [StructLayout(LayoutKind.Sequential)]
        struct CURSORINFO
        {
            public Int32 cbSize;        // Specifies the size, in bytes, of the structure. 
                                        // The caller must set this to Marshal.SizeOf(typeof(CURSORINFO)).
            public Int32 flags;         // Specifies the cursor state. This parameter can be one of the following values:
                                        //    0             The cursor is hidden.
                                        //    CURSOR_SHOWING    The cursor is showing.
            public IntPtr hCursor;          // Handle to the cursor. 
            public POINT ptScreenPos;       // A POINT structure that receives the screen coordinates of the cursor. 
        }

        [DllImport("user32.dll")]
        static extern bool GetCursorInfo(out CURSORINFO pci);



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
