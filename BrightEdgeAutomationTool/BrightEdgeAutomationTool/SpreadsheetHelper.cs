using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BrightEdgeAutomationTool
{
    public static class SpreadsheetHelper
    {
        public static WorkbookPart workbookPart;
        public static MainSheetData GetMainSheetData(WorksheetPart mainSheetPart)
        {
            MainSheetData replace = new MainSheetData();
            var rows = mainSheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();
            // Get Marsha
            var row = rows.ElementAtOrDefault<Row>(1);
            var marshaCell = GetRowCells(row).ElementAtOrDefault<Cell>(8);
            replace.Marsha = GetCellValue(marshaCell, workbookPart);

            row = rows.ElementAtOrDefault<Row>(3);
            var countryCell = GetRowCells(row).ElementAtOrDefault<Cell>(8);
            replace.Country = GetCellValue(countryCell, workbookPart);

            for(var i = 5; i < rows.Count(); i++)
            {
                row = rows.ElementAtOrDefault<Row>(i);
                var pageCell = GetRowCells(row).ElementAtOrDefault<Cell>(8);
                var pageValue = GetCellValue(pageCell, workbookPart);

                if(pageValue != null)
                {
                    if (!pageValue.ToString().Equals(""))
                        replace.Pages.Add(pageValue);
                }
                
            }

            return replace;
        }

        public static List<string> GetKeywordsFromSheet(string sheetName)
        {
            List<string> keywordList = new List<string>();
            var pagesSheetPart = GetWorksheetPartThatBeginsWith(workbookPart, sheetName.Substring(0, 3));
            var rows = pagesSheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            var keywordsRaw = rows.Select(kRow => GetCellValue(GetRowCells(kRow).ElementAtOrDefault(0), workbookPart));

            var keywordsNotNull = keywordsRaw.Where(c => {
                if (!string.IsNullOrEmpty(c))
                    return true;
                return false;
            });

            for (int i = 0; i < keywordsNotNull.Count(); i += 500) //should be 1000
            {
                //List<string> myNewString = tempSMSNoList.Skip(i * groupSize).Take(groupSize).ToList();
                var keywordListStr = keywordsNotNull.Skip(i == 0 ? 0 : i).Take(1000).Aggregate((x, y) => x + "\n" + y);
                keywordList.Add(keywordListStr);
            }

            return keywordList;
        }

        public static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            var resultSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => sheetName.Equals(s.Name));
            if (resultSheet == null)
                return null;

            string relId = resultSheet.Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }

        public static WorksheetPart GetWorksheetPartThatBeginsWith(WorkbookPart workbookPart, string partialSheetName)
        {
            var resultSheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name.ToString().StartsWith(partialSheetName,StringComparison.CurrentCultureIgnoreCase));
            if (resultSheet == null)
                return null;

            string relId = resultSheet.Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }






        // Retrieve the value of a cell, given a file name, sheet name,
        // and address name.
        public static string GetCellValue(Cell theCell, WorkbookPart wbPart)
        {
            string value = null;

            if (theCell == null)
                return value;

            // If the cell does not exist, return an empty string.
            if (theCell.InnerText.Length > 0)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;

                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                    }
                }
            }

            return value;
        }



        ///<summary>returns an empty cell when a blank cell is encountered
        ///</summary>
        public static IEnumerable<Cell> GetRowCells(Row row)
        {
            int currentCount = 0;

            foreach (DocumentFormat.OpenXml.Spreadsheet.Cell cell in
                row.Descendants<DocumentFormat.OpenXml.Spreadsheet.Cell>())
            {
                string columnName = GetColumnName(cell.CellReference);

                int currentColumnIndex = ConvertColumnNameToNumber(columnName);

                for (; currentCount < currentColumnIndex; currentCount++)
                {
                    yield return new DocumentFormat.OpenXml.Spreadsheet.Cell();
                }

                yield return cell;
                currentCount++;
            }
        }

        /// <summary>
        /// Given a cell name, parses the specified cell to get the column name.
        /// </summary>
        /// <param name="cellReference">Address of the cell (ie. B2)</param>
        /// <returns>Column Name (ie. B)</returns>
        public static string GetColumnName(string cellReference)
        {
            // Match the column name portion of the cell name.
            var regex = new System.Text.RegularExpressions.Regex("[A-Za-z]+");
            var match = regex.Match(cellReference);

            return match.Value;
        }

        /// <summary>
        /// Given just the column name (no row index),
        /// it will return the zero based column index.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful</returns>
        /// <exception cref="ArgumentException">thrown if the given string
        /// contains characters other than uppercase letters</exception>
        public static int ConvertColumnNameToNumber(string columnName)
        {
            var alpha = new System.Text.RegularExpressions.Regex("^[A-Z]+$");
            if (!alpha.IsMatch(columnName)) throw new ArgumentException();

            char[] colLetters = columnName.ToCharArray();
            Array.Reverse(colLetters);

            int convertedValue = 0;
            for (int i = 0; i < colLetters.Length; i++)
            {
                char letter = colLetters[i];
                int current = i == 0 ? letter - 65 : letter - 64; // ASCII 'A' = 65
                convertedValue += current * (int)Math.Pow(26, i);
            }

            return convertedValue;
        }
    }
}
