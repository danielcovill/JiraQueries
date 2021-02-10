using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using CsvHelper;

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
        public void GenerateTicketListSummary(JiraSearchResponse searchResponse, List<User> engineers, string xlsxOutputPath, string sourceQuery)
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

                var ticketsRow = ++rowIndex;
                InsertCellInWorksheet("A", rowIndex, maintTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("B", rowIndex, bugTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("C", rowIndex, taskTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("D", rowIndex, storyTicketCount.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("E", rowIndex, $"=SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("F", rowIndex, $"=A{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("G", rowIndex, $"=B{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("H", rowIndex, $"=C{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("I", rowIndex, $"=D{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);

                rowIndex += 2;
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

                var pointsRow = ++rowIndex;
                InsertCellInWorksheet("A", rowIndex, maintTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("B", rowIndex, bugTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("C", rowIndex, taskTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("D", rowIndex, storyTicketPts.ToString(), overviewSheetData, FormatOptions.Number);
                InsertCellInWorksheet("E", rowIndex, $"=SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("F", rowIndex, $"=A{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("G", rowIndex, $"=B{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("H", rowIndex, $"=C{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);
                InsertCellInWorksheet("I", rowIndex, $"=D{rowIndex}/SUM(A{rowIndex}:D{rowIndex})", overviewSheetData, FormatOptions.Percent, true);

                rowIndex += 2;
                InsertCellInWorksheet("A", rowIndex, "Averages", overviewSheetData, FormatOptions.String);

                rowIndex++;
                InsertCellInWorksheet("A", rowIndex, "Maint Avg", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("B", rowIndex, "Bug Avg", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("C", rowIndex, "Task Avg", overviewSheetData, FormatOptions.String);
                InsertCellInWorksheet("D", rowIndex, "Story Avg", overviewSheetData, FormatOptions.String);

                var averagesRow = ++rowIndex;
                InsertCellInWorksheet("A", rowIndex, $"=A{pointsRow}/A{ticketsRow}", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("B", rowIndex, $"=B{pointsRow}/B{ticketsRow}", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("C", rowIndex, $"=C{pointsRow}/C{ticketsRow}", overviewSheetData, FormatOptions.Number, true);
                InsertCellInWorksheet("D", rowIndex, $"=D{pointsRow}/D{ticketsRow}", overviewSheetData, FormatOptions.Number, true);

                //Add a pie chart showing the ratio of tickets by type (ttr = TicketTypeRatio)
                var chartTitle = "Ticket Type Ratio";

                //cell reference, value
                DrawingsPart ttrDrawingPart = rollupWorksheet.WorksheetPart.AddNewPart<DrawingsPart>();
                var rollupWorksheetPart = rollupWorksheet.WorksheetPart;
                rollupWorksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
                {
                    Id = rollupWorksheet.WorksheetPart.GetIdOfPart(ttrDrawingPart)
                });
                rollupWorksheetPart.Worksheet.Save();
                ChartPart ttrChartPart = ttrDrawingPart.AddNewPart<ChartPart>();
                ttrChartPart.ChartSpace = new ChartSpace();
                ttrChartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });

                Chart ttrChart = ttrChartPart.ChartSpace.AppendChild<Chart>(new Chart());
                PlotArea ttrPlotArea = ttrChart.AppendChild<PlotArea>(new PlotArea());
                Layout ttrLayout = ttrPlotArea.AppendChild<Layout>(new Layout());
                PieChart ttrPieChart = ttrPlotArea.AppendChild<PieChart>(new PieChart(/*Maybe something needs to be here?*/));
                PieChartSeries ttrPieChartSeries = ttrPieChart.AppendChild<PieChartSeries>(
                    new PieChartSeries(
                        new DocumentFormat.OpenXml.Drawing.Charts.Index()
                        {
                            Val = 1U
                        }
                    )
                );


                // Iterate through the data we want in the pie chart
                // StringLiteral strLit = ttrPieChartSeries.AppendChild<CategoryAxisData>(
                //     new CategoryAxisData()).AppendChild<StringLiteral>(new StringLiteral()
                // );
                // strLit.Append(new PointCount() { Val = 1U });
                // strLit.AppendChild<StringPoint>(new StringPoint() 
                // { 
                //     Index = 0U 
                // }).Append(new NumericValue(chartTitle));
                // NumberReference numRef = ttrPieChartSeries
                //     .AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(new DocumentFormat.OpenXml.Drawing.Charts.Values())
                //     .AppendChild<NumberReference>(new NumberReference());
                // numRef.Append(new FormatCode("General"));
                // numRef.Append(new PointCount() { Val = 1U });
                // numRef.AppendChild<NumericPoint>(new NumericPoint() { Index = 0U }).Append(new NumericValue(data[key].ToString()));

                //Add a pie chart showing the ratio of points by type


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
                    InsertCellInWorksheet("B", rowIndex, $"=SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Number, true);
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
                    InsertCellInWorksheet("B", rowIndex, $"=SUM(C{rowIndex}:F{rowIndex})", devBreakdownSheetData, FormatOptions.Number, true);
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

                // Create sheet with details of query
                var metaDataWorksheet = CreateSheetInDocument(spreadsheetDocument, "metadata");
                var metaDataSheetData = metaDataWorksheet.GetFirstChild<SheetData>();
                InsertCellInWorksheet("A", 1, "JQL", metaDataSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 2, sourceQuery, metaDataSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 4, "Run date/time", metaDataSheetData, FormatOptions.String);
                InsertCellInWorksheet("A", 5, DateTime.Now.ToString(), metaDataSheetData, FormatOptions.String);

            }
        }

        public void GenerateTeamOutputReport(JiraSearchResponse searchResponse, List<User> engineers, string csvOutputPath)
        {
            var results = new List<EngineerOutputReport>();
            foreach (var engineer in engineers)
            {
                results.Add(new EngineerOutputReport()
                {
                    Name = engineer.displayName,
                    TenWeekAveragePoints = searchResponse.GetPointTotal(null, engineer.accountId, DateTime.Now.AddDays(-70), DateTime.Now) / 10,
                    TwoWeekAveragePoints = searchResponse.GetPointTotal(null, engineer.accountId, DateTime.Now.AddDays(-14), DateTime.Now) / 2,
                    PriorWeekPoints = searchResponse.GetPointTotal(null, engineer.accountId, DateTime.Now.AddDays(-7), DateTime.Now)
                });
            }
            using (var writer = new StreamWriter(csvOutputPath))
            using (var csv = new CsvWriter(writer, CultureInfo.InvariantCulture))
            {
                csv.WriteRecords(results);
            }
        }

        public void GenerateTimespanBreakdown(JiraSearchResponse searchResponse, List<User> engineers, DateTime startDate, DateTime endDate, string xlsxOutputPath)
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
            if (sheets == null)
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