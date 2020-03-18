using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Shapes;

namespace BrightEdgeAutomationTool
{
    public static class RTSetup
    {
        public static MainWindow MainWindow;
        public static HotKeyForm HotKeyForm;
        private delegate void SetCoords();
        private static int Index = 0;
        private const string HOTKEYS = "SHIFT + A";
        private const string HOTKEY_CANCEL = "SHIFT + C";
        public static RTSettings rtSettings;

        public enum RTProcess
        {
            Run,
            Stop
        }

        private static List<SetCoords> CoordProcesses = new List<SetCoords>()
        {
            Initialize,
            AddRTDimensions,
            AddKeywordsCoords,
            AddKeywordsInputCoords,
            NextButtonCoords,
            FinishButtonCoords,
            ProgressbarCoords,
            DownloadCoords,
            TableTopCoords,
            StopTaskCoords,
            StopTaskApplyCoords,
            Finalizer
        };

        private static List<string> Instructions = new List<string>()
        {
            "",
            $"Hit {HOTKEYS} to save Rank Tracker dimensions and position",
            $"Hit {HOTKEYS} to set the 'Add Keywords' button position",
            $"Hit {HOTKEYS} to set the Keyword input area position",
            $"Hit {HOTKEYS} to set the 'Next' button position",
            $"Hit {HOTKEYS} to set the 'Finish' button position",
            $"Hit {HOTKEYS} to set the 'progressbar' position",
            $"Hit {HOTKEYS} to set the 'Download' button position",
            $"Hit {HOTKEYS} to set the point of the top of the keyword list table",
            $"Hit {HOTKEYS} to set the Stop Task Button position",
            $"Hit {HOTKEYS} to set the Stop Task Apply Button position",
            $"Hit {HOTKEYS} to save settings"
        };

        public static void RunProcess(RTProcess rtp)
        {
            if(rtp == RTProcess.Run)
            {
                CoordProcesses.ElementAt(Index)();
            }

            if (rtp == RTProcess.Stop)
            {
                Index = 0;
                MainWindow.UpdateStatus("Process stopped");
                HotKeyForm.UnRegisterHotKey();
            }
        }

        private static void DisplayNextInstructions()
        {
            MainWindow.UpdateStatus(Instructions[Index]);
        }

        private static void Initialize()
        {
            rtSettings = new RTSettings();
            MainWindow.UpdateStatus($"Starting the Update Rank Tracker Process");
            MainWindow.UpdateStatus($"Hit {HOTKEY_CANCEL} to cancel");
            MainWindow.UpdateStatus($"--------------------------------------------");
            
            Index += 1;
            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to save Rank Tracker dimensions and position");
            DisplayNextInstructions();
        }

        private static void AddRTDimensions()
        {
            var rect = HWNDHelper.GetRankTrackerRect();
            MainWindow.UpdateStatus($"Rank Tracker: X:{rect.X} Y:{rect.Y} Width:{rect.Width} Height:{rect.Height}");
            rtSettings.RankTrackerRect = rect;
            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Add Keywords' button position");
            DisplayNextInstructions();
        }

        private static void AddKeywordsCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Add Keyword button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.AddKeywordPos = cPosition;
            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the Keyword input area position");
            DisplayNextInstructions();
        }

        private static void AddKeywordsInputCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Keyword Input Area: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.KeywordInputPos = cPosition;
            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Next' button position");
            DisplayNextInstructions();
        }

        private static void NextButtonCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Next button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.NextPos = cPosition;
            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Finish' button position");
            DisplayNextInstructions();
        }

        private static void FinishButtonCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Finish button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.FinishPos = cPosition;
            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'progressbar' position");
            DisplayNextInstructions();
        }

        private static void ProgressbarCoords()
        {
            string pGreenColor = "#67B847"; // Progressbar green color
            string pGreenColor2 = "#61A032"; // Progressbar green color
            string pDarkColor = "#1F2530"; // Progressbar dark color

            var cPosition = Cursor.Position;

            // Check if point x & y matches colors of the progressbar

            var c = HexConverter(RankTrackerPuppetMaster.GetColorAt(cPosition.X, cPosition.Y));

            if (!c.Equals(pGreenColor) && !c.Equals(pDarkColor) && !c.Equals(pGreenColor2))
            {
                MainWindow.UpdateStatus($"Unable to detect the progress bar based on the color of the selected pixel");
                return;
            }

            MainWindow.UpdateStatus($"Progress Bar at: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.ProgressbarPos = cPosition;
            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Download' button position");
            DisplayNextInstructions();
        }

        private static void DownloadCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Download button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.DownloadPos = cPosition;

            Index += 1;

            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the point of the top of the keyword list table");
            DisplayNextInstructions();
        }


        private static void TableTopCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Keyword table: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.KeywordTblPos = cPosition;
            Index += 1;
            //MainWindow.UpdateStatus($"Hit {HOTKEYS} to save settings");
            DisplayNextInstructions();
        }

        private static void StopTaskCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Stop Task Button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.TaskStopPos = cPosition;
            Index += 1;

            DisplayNextInstructions();
        }

        private static void StopTaskApplyCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Stop Task Apply Button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.TaskStopApplyPos = cPosition;
            Index += 1;

            DisplayNextInstructions();
        }


        private static void Finalizer()
        {
            Index = 0;
            var result = MainWindow.database.UpdateRTSettings(rtSettings);
            MainWindow.UpdateStatus($"Process complete");
            HotKeyForm.UnRegisterHotKey();

            //var sets = MainWindow.database.GetRTSettings();
            //Console.WriteLine($"Rank Tracker: X:{sets.RankTrackerRect.X} Y:{sets.RankTrackerRect.Y} Width:{sets.RankTrackerRect.Width} Height:{sets.RankTrackerRect.Height}");
        }



        private static String HexConverter(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

    }

    [Serializable]
    public class RTSettings
    {
        public System.Drawing.Rectangle RankTrackerRect { get; set; }
        public Point AddKeywordPos { get; set; }
        public Point KeywordInputPos { get; set; }
        public Point NextPos { get; set; }
        public Point FinishPos { get; set; }
        public Point ProgressbarPos { get; set; }
        public Point DownloadPos { get; set; }
        public Point CloseUpgradePos { get; set; }
        public Point KeywordTblPos { get; set; }
        public Point TaskStopPos { get; set; }
        public Point TaskStopApplyPos { get; set; }
    }
}
