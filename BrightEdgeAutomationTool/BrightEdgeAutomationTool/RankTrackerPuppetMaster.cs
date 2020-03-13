using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BrightEdgeAutomationTool
{
    public static class RankTrackerPuppetMaster
    {
        public static void StartProcess(DirectoryInfo directory)
        {
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

                            foreach (var res in distinctResultSheetData)
                            {
                                Console.WriteLine($"{res.Keyword} - {res.Volume}");
                            }
                        }
                    }
                }
            }
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
