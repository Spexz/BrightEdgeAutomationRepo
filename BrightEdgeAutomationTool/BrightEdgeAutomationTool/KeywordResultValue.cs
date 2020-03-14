using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BrightEdgeAutomationTool
{
    public class KeywordResultValue
    {
        public string Keyword;
        public decimal Volume;
        public string GoogleRank;
        public string RankingPage;

        public static KeywordResultValue FromCsv(string csvLine)
        {
            string sep = ",";

            KeywordResultValue kValues = null;
            string[] values = csvLine.Split(sep.ToCharArray());

            if (Convert.ToDecimal(values[1]) == 0)
                return kValues;


            kValues = new KeywordResultValue();

            kValues.Keyword = values[0];
            kValues.Volume = Convert.ToDecimal(values[1]);

            return kValues;
        }

        public static KeywordResultValue FromRankTrackerCsv(string csvLine)
        {
            string sep = ",";

            KeywordResultValue kValues = null;
            string[] values = csvLine.Split(sep.ToCharArray());

            kValues = new KeywordResultValue();

            kValues.Keyword = values[0];
            kValues.GoogleRank = values[2];
            kValues.RankingPage = values[3];

            return kValues;
        }
    }
}
