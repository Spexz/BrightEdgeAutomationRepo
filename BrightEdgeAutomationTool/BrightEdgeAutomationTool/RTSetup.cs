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
        private delegate void SetCoords();
        private static int Index = 0;
        private const string HOTKEYS = "SHIFT + A";
        public static RTSettings rtSettings;

        private static List<SetCoords> CoordProcesses = new List<SetCoords>()
        {
            Initialize,
            AddRTDimensions,
            AddKeywordsCoords,
            NextButtonCoords,
            FinishButtonCoords,
            ProgressbarCoords,
            DownloadCoords,
            TableTopCoords
        };

        public static void RunProcess()
        {
            CoordProcesses.ElementAt(Index)();
        }

        private static void Initialize()
        {
            rtSettings = new RTSettings();
            MainWindow.UpdateStatus($"Hit {HOTKEYS} to save Rank Tracker dimensions and postion");
            Index += 1;
        }

        private static void AddRTDimensions()
        {
            var rect = HWNDHelper.GetRankTrackerRect();
            MainWindow.UpdateStatus($"Rank Tracker: X:{rect.X} Y:{rect.Y} Width:{rect.Width} Height:{rect.Height}");
            rtSettings.RankTrackerRect = rect;
            Index += 1;

            MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Add Keywords' button position");
        }

        private static void AddKeywordsCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Add Keyword button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.AddKeywordPos = cPosition;
            Index += 1;

            MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Next' button position");
        }

        private static void NextButtonCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Next button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.NextPos = cPosition;
            Index += 1;

            MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Finish' button position");
        }

        private static void FinishButtonCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Finish button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.FinishPos = cPosition;
            Index += 1;

            MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'progressbar' position");
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
                MainWindow.UpdateStatus($"Unable to detect the progressbar based on the color of the selected pixel");
                return;
            }

            MainWindow.UpdateStatus($"Progressbar at: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.ProgressbarPos = cPosition;
            Index += 1;

            MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the 'Download' button position");
        }

        private static void DownloadCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Download button: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.DownloadPos = cPosition;
            Index += 1;

            MainWindow.UpdateStatus($"Hit {HOTKEYS} to set the point of the top of the keyword list table");
        }


        private static void TableTopCoords()
        {
            var cPosition = Cursor.Position;
            MainWindow.UpdateStatus($"Keyword table: X: {cPosition.X} Y: {cPosition.Y}");
            rtSettings.KeywordTblPos = cPosition;
            Index += 1;

        }

        private static void Finalize()
        {
            Index = 0;
            var result = MainWindow.database.UpdateRTSettings(rtSettings);
            MainWindow.UpdateStatus($"Process complete");
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
        public Point NextPos { get; set; }
        public Point FinishPos { get; set; }
        public Point ProgressbarPos { get; set; }
        public Point DownloadPos { get; set; }
        public Point KeywordTblPos { get; set; }
    }
}
