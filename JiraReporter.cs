using System;
using System.Collections.Generic;
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
        public void GenerateTicketListSummary(JiraSearchResponse searchResponse, List<User> engineers, string xlsxOutputPath)
        {
            if (xlsxOutputPath == null)
            {
                throw new Exception("Output path required");
            }
            if (File.Exists(xlsxOutputPath))
            {
                throw new Exception("File already exists");
            }
            using (var spreadsheetDocument = CreateEmptySpreadsheet(xlsxOutputPath))
            {
                //Create Rollup Worksheet
                var rollupWorksheet = CreateSheetInDocument(spreadsheetDocument, "Rollup");
                var overviewSheetData = rollupWorksheet.GetFirstChild<SheetData>();

                uint rowIndex = 1;
                var totalTickets = searchResponse.GetTicketCount();
                var totalPts = searchResponse.GetPointTotal();
                var maintTicketCount = searchResponse.GetTicketCount(maintenance);
                var maintTicketPts = searchResponse.GetPointTotal(maintenance);
                var bugTicketCount = searchResponse.GetTicketCount(bug);
                var bugTicketPts = searchResponse.GetPointTotal(bug);
                var taskTicketCount = searchResponse.GetTicketCount(task);
                var taskTicketPts = searchResponse.GetPointTotal(task);
                var storyTicketCount = searchResponse.GetTicketCount(story);
                var storyTicketPts = searchResponse.GetPointTotal(story);

                InsertCellInWorksheet("A", rowIndex, "Tickets", overviewSheetData, FormatOptions.String);

                rowIndex++;
                InsertCellInWorksheet("A", rowIndex, "Maint Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", rowIndex, "Bug Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("C", rowIndex, "Task Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("D", rowIndex, "Story Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("E", rowIndex, "Total Tkts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("F", rowIndex, "Maint %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("G", rowIndex, "Bug %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("H", rowIndex, "Task %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("I", rowIndex, "Story %", overviewSheetData, FormatOptions.String);
                rowIndex++;

                InsertCellInWorksheet("A", rowIndex, maintTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("B", rowIndex, bugTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("C", rowIndex, taskTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("D", rowIndex, storyTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("E", rowIndex, totalTickets.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("F", rowIndex, "=A6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("G", rowIndex, "=B6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("H", rowIndex, "=C6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("I", rowIndex, "=D6/SUM(A6:D6)", overviewSheetData, FormatOptions.Percent, true);
                rowIndex+=2;

                InsertCellInWorksheet("A", rowIndex, "Points", overviewSheetData, FormatOptions.String);
                rowIndex++;

                InsertCellInWorksheet("A", rowIndex, "Maint Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", rowIndex, "Bug Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("C", rowIndex, "Task Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("D", rowIndex, "Story Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("E", rowIndex, "Total Pts", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("F", rowIndex, "Maint %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("G", rowIndex, "Bug %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("H", rowIndex, "Task %", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("I", rowIndex, "Story %", overviewSheetData, FormatOptions.String);
                rowIndex++;

                InsertCellInWorksheet("A", rowIndex, maintTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("B", rowIndex, bugTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("C", rowIndex, taskTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("D", rowIndex, storyTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("E", rowIndex, totalPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("F", rowIndex, "=A10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("G", rowIndex, "=B10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("H", rowIndex, "=C10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("I", rowIndex, "=D10/SUM(A10:D10)", overviewSheetData, FormatOptions.Percent, true);
                rowIndex+=2;

                InsertCellInWorksheet("A", rowIndex, "Averages", overviewSheetData, FormatOptions.String);
                rowIndex++;

                InsertCellInWorksheet("A", rowIndex, "Maint Avg", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", rowIndex, "Bug Avg", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("C", rowIndex, "Task Avg", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("D", rowIndex, "Story Avg", overviewSheetData, FormatOptions.String);
                rowIndex++;

                InsertCellInWorksheet("A", rowIndex, "=A10/A6", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("B", rowIndex, "=B10/B6", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("C", rowIndex, "=C10/C6", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("D", rowIndex, "=D10/D6", overviewSheetData, FormatOptions.Number, true);
                //TODO: Fix the formulas because the rows are all different now that I pulled the date fields

                //Create Per Developer Breakdown Worksheet
                rowIndex = 1;
                var devBreakdownWorksheet = CreateSheetInDocument(spreadsheetDocument, "DevBreakdown");
                var devBreakdownSheetData = devBreakdownWorksheet.GetFirstChild<SheetData>();

                foreach (var engineer in engineers)
                {
                    rowIndex++;
                    InsertCellInWorksheet("A", rowIndex, engineer.displayName, devBreakdownSheetData, FormatOptions.String);
                    rowIndex++;
                    //Ticket Count Header Row
                    InsertCellInWorksheet("B", rowIndex, "Total Tkts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("C", rowIndex, "Maint Tkts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("D", rowIndex, "Bug Tkts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("E", rowIndex, "Task Tkts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("F", rowIndex, "Story Tkts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("G", rowIndex, "Maint %", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("H", rowIndex, "Bug %", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("I", rowIndex, "Task %", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("J", rowIndex, "Story %", devBreakdownSheetData, FormatOptions.String);
                    rowIndex++;
                    //Ticket Counts
                    InsertCellInWorksheet("B", rowIndex, searchResponse.GetTicketCount(accountId: engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("C", rowIndex, searchResponse.GetTicketCount(maintenance, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("D", rowIndex, searchResponse.GetTicketCount(bug, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("E", rowIndex, searchResponse.GetTicketCount(task, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("F", rowIndex, searchResponse.GetTicketCount(story, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    //Ticket Percentages
                    InsertCellInWorksheet("G", rowIndex, $"=C{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    InsertCellInWorksheet("H", rowIndex, $"=D{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    InsertCellInWorksheet("I", rowIndex, $"=E{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    InsertCellInWorksheet("J", rowIndex, $"=F{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    rowIndex++;
                    //Points Header Row
                    InsertCellInWorksheet("B", rowIndex, "Total Pts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("C", rowIndex, "Maint Pts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("D", rowIndex, "Bug Pts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("E", rowIndex, "Task Pts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("F", rowIndex, "Story Pts", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("G", rowIndex, "Maint %", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("H", rowIndex, "Bug %", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("I", rowIndex, "Task %", devBreakdownSheetData, FormatOptions.String);
                    InsertCellInWorksheet("J", rowIndex, "Story %", devBreakdownSheetData, FormatOptions.String);
                    rowIndex++;
                    //Points
                    InsertCellInWorksheet("B", rowIndex, searchResponse.GetPointTotal(accountId: engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("C", rowIndex, searchResponse.GetPointTotal(maintenance, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("D", rowIndex, searchResponse.GetPointTotal(bug, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("E", rowIndex, searchResponse.GetPointTotal(task, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    InsertCellInWorksheet("F", rowIndex, searchResponse.GetPointTotal(story, engineer.accountId).ToString(), devBreakdownSheetData, FormatOptions.Number);
                    //Points Percentages
                    InsertCellInWorksheet("G", rowIndex, $"=C{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    InsertCellInWorksheet("H", rowIndex, $"=D{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    InsertCellInWorksheet("I", rowIndex, $"=E{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    InsertCellInWorksheet("J", rowIndex, $"=F{rowIndex}/SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Percent, true);
                    rowIndex++;
                }
            }
            //TODO: Add sheet with the query and date run at the end
        }

        public void GenerateWeeklySummary(JiraSearchResponse searchResponse, List<User> engineers, DateTime startDate, DateTime endDate, string xlsxOutputPath) 
        {
            throw new NotImplementedException();
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