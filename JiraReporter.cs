using System;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace work_charts
{

    class JiraReporter
    {

        private enum FormatOptions
        {
            String,
            Number,
            Date,
            DateTime,
            Percent
        };

        private const string maintenance = "Maintenance";
        private const string bug = "Bug";
        private const string task = "Task";
        private const string story = "Story";
        public JiraReporter()
        {

        }
        /// <summary>
        /// Gets a breakdown of ticket counts, points, and ratios among ticket types for a given list of tickets
        /// </summary>
        /// <param name="issues">The list of issues over which to compute</param>
        /// <param name="xlsxOutputPath">If included output creates an xlsx document at the specified path. Otherwise output goes to the console.</param>
        public void GenerateTicketListSummary(JiraSearchResponse searchResponse, DateTime startDate, DateTime endDate, string xlsxOutputPath = null)
        {
            if (xlsxOutputPath == null)
            {
                Console.WriteLine("-- Ticket Counts --");
                Console.WriteLine($"Maintenance Tickets: {searchResponse.GetTicketCount(maintenance)}");
                Console.WriteLine($"Bug Tickets        : {searchResponse.GetTicketCount(bug)}");
                Console.WriteLine($"Task Tickets       : {searchResponse.GetTicketCount(task)}");
                Console.WriteLine($"Story Tickets      : {searchResponse.GetTicketCount(story)}");
                Console.WriteLine();

                Console.WriteLine("-- Ticket Ratios --");
                double sumTotalTickets = searchResponse.GetTicketCount();
                Console.WriteLine($"Maintenance Tickets: {(searchResponse.GetTicketCount(maintenance) / sumTotalTickets).ToString("P")}");
                Console.WriteLine($"Bug Tickets        : {(searchResponse.GetTicketCount(bug) / sumTotalTickets).ToString("P")}");
                Console.WriteLine($"Task Tickets       : {(searchResponse.GetTicketCount(task) / sumTotalTickets).ToString("P")}");
                Console.WriteLine($"Story Tickets      : {(searchResponse.GetTicketCount(story) / sumTotalTickets).ToString("P")}");
                Console.WriteLine();

                Console.WriteLine("-- Point Totals --");
                Console.WriteLine($"Maintenance Points: {searchResponse.GetPointTotal(maintenance)}");
                Console.WriteLine($"Bug Points        : {searchResponse.GetPointTotal(bug)}");
                Console.WriteLine($"Task Points       : {searchResponse.GetPointTotal(task)}");
                Console.WriteLine($"Story Points      : {searchResponse.GetPointTotal(story)}");
                Console.WriteLine();

                Console.WriteLine("-- Point Ratios --");
                double sumTotalPoints = searchResponse.GetPointTotal();
                Console.WriteLine($"Maintenance Points: {(searchResponse.GetPointTotal(maintenance) / sumTotalPoints).ToString("P")}");
                Console.WriteLine($"Bug Points        : {(searchResponse.GetPointTotal(bug) / sumTotalPoints).ToString("P")}");
                Console.WriteLine($"Task Points       : {(searchResponse.GetPointTotal(task) / sumTotalPoints).ToString("P")}");
                Console.WriteLine($"Story Points      : {(searchResponse.GetPointTotal(story) / sumTotalPoints).ToString("P")}");
                Console.WriteLine();

                Console.WriteLine("-- Per Employee Stats --");
                foreach (var assigneeIssuesSet in searchResponse.issues.GroupBy(issue => issue.fields.assignee))
                {
                    if (assigneeIssuesSet.Key == null)
                    {
                        Console.WriteLine("- Unassigned");
                    }
                    else
                    {
                        Console.WriteLine($"- {assigneeIssuesSet.Key.displayName}");
                    }
                    Console.WriteLine($"Total Tickets      : {assigneeIssuesSet.Count()}");
                    Console.WriteLine($"Maintenance Tickets: {assigneeIssuesSet.Count(issue => issue.fields.issuetype.name == maintenance)}");
                    Console.WriteLine($"Bug Tickets        : {assigneeIssuesSet.Count(issue => issue.fields.issuetype.name == bug)}");
                    Console.WriteLine($"Task Tickets       : {assigneeIssuesSet.Count(issue => issue.fields.issuetype.name == task)}");
                    Console.WriteLine($"Story Tickets      : {assigneeIssuesSet.Count(issue => issue.fields.issuetype.name == story)}");
                    Console.WriteLine();
                }
                return;
            }

            if (File.Exists(xlsxOutputPath))
            {
                throw new Exception("File already exists");
            }
            // SpreadsheetDocument 
            // -> WorkbookPart 
            //     -> Workbook
            //         -> <Sheets> -> Sheet
            //     -> WorksheetPart 
            //         -> Worksheet 
            //             -> SheetData
            using (var spreadsheetDocument = CreateEmptySpreadsheet(xlsxOutputPath))
            {
                uint sheetIdIterator = 1;
                var workbookPart = spreadsheetDocument.WorkbookPart;

                var overviewWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var overviewSheetData = new SheetData();
                overviewWorksheetPart.Worksheet = new Worksheet(overviewSheetData);

                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
                var summarySheet = new Sheet()
                {
                    Name = "Rollup",
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(overviewWorksheetPart),
                    SheetId = sheetIdIterator++
                };
                sheets.Append(summarySheet);

                InsertCellInWorksheet("A", 4, "Tickets", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 8, "Points", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 12, "Averages", overviewSheetData, FormatOptions.String);

                InsertCellInWorksheet("E", 5, "Total Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("E", 9, "Total Pts", overviewSheetData, FormatOptions.String);
                var totalTickets = searchResponse.GetTicketCount();
                var totalPts = searchResponse.GetPointTotal();
                InsertCellInWorksheet("E", 6, totalTickets.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("E", 10, totalPts.ToString(), overviewSheetData, FormatOptions.Number);

                InsertCellInWorksheet("A", 1, "Start Date", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 2, startDate.ToShortDateString(), overviewSheetData, FormatOptions.Date);
                InsertCellInWorksheet("B", 1, "End Date", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", 2, endDate.ToShortDateString(), overviewSheetData, FormatOptions.Date);

                InsertCellInWorksheet("A", 5, "Maint Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 9, "Maint Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 13, "Maint Avg", overviewSheetData, FormatOptions.String);
                var maintTicketCount = searchResponse.GetTicketCount(maintenance);
                var maintTicketPts = searchResponse.GetPointTotal(maintenance);
                InsertCellInWorksheet("A", 6, maintTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("A", 10, maintTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("A", 14, "=A10/A6", overviewSheetData, FormatOptions.Number, true);


                InsertCellInWorksheet("F", 5, "Maint %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("F", 9, "Maint %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("F", 6, "=A6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("F", 10, "=A10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);

                InsertCellInWorksheet("B", 5, "Bug Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", 9, "Bug Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", 13, "Bug Avg", overviewSheetData, FormatOptions.String);
                var bugTicketCount = searchResponse.GetTicketCount(bug);
                var bugTicketPts = searchResponse.GetPointTotal(bug);
                InsertCellInWorksheet("B", 6, bugTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("B", 10, bugTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("B", 14, "=B10/B6", overviewSheetData, FormatOptions.Number, true);

                InsertCellInWorksheet("G", 5, "Bug %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("G", 9, "Bug %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("G", 6, "=B6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("G", 10, "=B10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);

                InsertCellInWorksheet("C", 5, "Task Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("C", 9, "Task Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("C", 13, "Task Avg", overviewSheetData, FormatOptions.String);
                var taskTicketCount = searchResponse.GetTicketCount(task);
                var taskTicketPts = searchResponse.GetPointTotal(task);
                InsertCellInWorksheet("C", 6, taskTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("C", 10, taskTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("C", 14, "=C10/C6", overviewSheetData, FormatOptions.Number, true);

                InsertCellInWorksheet("H", 5, "Task %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("H", 9, "Task %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("H", 6, "=C6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("H", 10, "=C10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);

                InsertCellInWorksheet("D", 5, "Story Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("D", 9, "Story Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("D", 13, "Story Avg", overviewSheetData, FormatOptions.String);
                var storyTicketCount = searchResponse.GetTicketCount(story);
                var storyTicketPts = searchResponse.GetPointTotal(story);
                InsertCellInWorksheet("D", 6, storyTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("D", 10, storyTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("D", 14, "=D10/D6", overviewSheetData, FormatOptions.Number, true);

                InsertCellInWorksheet("I", 5, "Story %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("I", 9, "Story %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("I", 6, "=D6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("I", 10, "=D10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);

                overviewWorksheetPart.Worksheet.Save();


                var devWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
                var devSheetData = new SheetData();
                devWorksheetPart.Worksheet = new Worksheet(devSheetData);

                var devBreakdownSheet = new Sheet()
                {
                    Name = "DevBreakdown",
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(devWorksheetPart),
                    SheetId = sheetIdIterator++
                };
                sheets.Append(devBreakdownSheet);
                workbookPart.Workbook.Save();
            }
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, string cellValue, SheetData sheetData, FormatOptions format, bool isFormula = false)
        {
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            // and make sure to do it in the right order...fkn microsoft example is broken
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                Row refRow = null;
                foreach (var rowIterater in sheetData.Elements<Row>())
                {
                    if (rowIterater.RowIndex > rowIndex)
                    {
                        refRow = rowIterater;
                        break;
                    }
                }
                row = new Row() { RowIndex = rowIndex };
                sheetData.InsertBefore(row, refRow);
            }

            // If there is not a cell with the specified column name, insert one and set it's value.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).Count() > 0)
            {
                var newCell = row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
                newCell.CellValue = new CellValue(cellValue);
                newCell.DataType = CellValues.String;
                return newCell;
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference.Value.Length == cellReference.Length)
                    {
                        if (String.Compare(cell.CellReference.Value, cellReference, true) > 0)
                        {
                            refCell = cell;
                            break;
                        }
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                //set data type
                if (isFormula)
                {
                    newCell.CellFormula = new CellFormula(cellValue);
                }
                else
                {
                    switch (format)
                    {
                        case FormatOptions.DateTime:
                        case FormatOptions.Date:
                            newCell.CellValue = new CellValue(DateTime.Parse(cellValue.ToString()).ToOADate().ToString(CultureInfo.InvariantCulture));//ew
                            break;
                        case FormatOptions.Number:
                        case FormatOptions.Percent:
                        case FormatOptions.String:
                            newCell.CellValue = new CellValue(cellValue);
                            break;
                    }
                }

                //set format
                switch (format)
                {
                    case FormatOptions.DateTime:
                        newCell.StyleIndex = 2;
                        newCell.DataType = new EnumValue<CellValues>(CellValues.Number);//surprisingly not date, date is only in Excel 2010
                        break;
                    case FormatOptions.Date:
                        newCell.StyleIndex = 1;
                        newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                        break;
                    case FormatOptions.String:
                        newCell.StyleIndex = 0;
                        newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                        break;
                    case FormatOptions.Number:
                        newCell.StyleIndex = 0;
                        break;
                    case FormatOptions.Percent:
                        newCell.StyleIndex = 3;
                        break;
                    default:
                        throw new Exception("Unexpected format option selected.");
                }
                row.InsertBefore(newCell, refCell);

                return newCell;
            }
        }

        private static Worksheet CreateSheetInDocument(SpreadsheetDocument spreadsheetDocument, string sheetName)
        {
            var workbookPart = spreadsheetDocument.WorkbookPart;
            var worksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            //add a sheets section if one doesn't already exist
            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            if(sheets == null)
            {
                sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            }

            var sheetIdIterator = Convert.ToUInt32(sheets.ChildElements.Count()) + 1;
            var summarySheet = new Sheet()
            {
                Name = sheetName,
                Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = sheetIdIterator
            };
            sheets.Append(summarySheet);

            return worksheetPart.Worksheet;
        }

        private static SpreadsheetDocument CreateEmptySpreadsheet(string xlsxOutputPath)
        {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(xlsxOutputPath, SpreadsheetDocumentType.Workbook);
            var workbookPart = spreadsheetDocument.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            // Add minimal Stylesheet
            var stylesPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet
            {
                Fonts = new Fonts(new Font()),
                Fills = new Fills(new Fill()),
                Borders = new Borders(new Border()),
                CellStyleFormats = new CellStyleFormats(new CellFormat()),
                CellFormats = new CellFormats(
                    new CellFormat(),
                    new CellFormat
                    {
                        NumberFormatId = 14,//format for date
                        ApplyNumberFormat = true
                    },
                    new CellFormat
                    {
                        NumberFormatId = 22,//format for datetime
                        ApplyNumberFormat = true
                    },
                    new CellFormat
                    {
                        NumberFormatId = 10,//format for percent
                        ApplyNumberFormat = true
                    }
                )
            };

            return spreadsheetDocument;
        }
    }
}