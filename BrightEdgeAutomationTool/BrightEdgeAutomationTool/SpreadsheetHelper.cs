using DocumentFormat.OpenXml;
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
                {
                    if (c.Equals("0"))
                        return false;
                    return true;
                }
                   
                return false;
            });

            for (int i = 0; i < keywordsNotNull.Count(); i += 1000) //should be 1000
            {
                //List<string> myNewString = tempSMSNoList.Skip(i * groupSize).Take(groupSize).ToList();
                var keywordListStr = keywordsNotNull.Skip(i == 0 ? 0 : i).Take(1000).Aggregate((x, y) => x + "\n" + y);
                keywordList.Add(keywordListStr);
            }

            return keywordList;
        }




        public static bool CreateResultSheet(List<KeywordResultValue> keywordStats)
        {
            var deleteResult = DeleteWorksheetPart("Results");


            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);


            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Results";

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };

            //sheets.Append(sheet);
            sheets.InsertAt(sheet, 1);
            // Save the new worksheet.
            newWorksheetPart.Worksheet.Save();


            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (workbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            }


            for (uint i = 0; i < keywordStats.Count(); i++)
            {
                Cell statB = InsertCellInWorksheet("B", i + 2, newWorksheetPart);
                Cell statC = InsertCellInWorksheet("C", i + 2, newWorksheetPart);

                // Insert the text into the SharedStringTablePart.
                var index = InsertSharedStringItem(keywordStats.ElementAt((int)i).Keyword, shareStringPart);
                // Set the value of cell B*
                statB.CellValue = new CellValue(index.ToString());
                statB.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                
                statC.CellValue = new CellValue(keywordStats.ElementAt((int)i).Volume.ToString());
                statC.DataType = new EnumValue<CellValues>(CellValues.Number);
            }


            // Save the new worksheet.
            newWorksheetPart.Worksheet.Save();



            return true;
        }

        public static bool DeleteWorksheetPart(string sheetName)
        {
            var resultSheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => sheetName.Equals(s.Name));
            if (resultSheet == null)
                return false;

            string relId = resultSheet.Id;

            // Remove the sheet reference from the workbook.
            var worksheetPart = (WorksheetPart)(workbookPart.GetPartById(relId));
            resultSheet.Remove();

            // Delete the worksheet part.
            workbookPart.DeletePart(worksheetPart);
            return true;
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


        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
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
