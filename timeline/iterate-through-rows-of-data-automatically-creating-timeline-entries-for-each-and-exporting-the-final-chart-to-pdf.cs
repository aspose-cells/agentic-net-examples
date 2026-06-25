using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Timelines;

class TimelineChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // -------------------------------------------------
            // 1. Populate sample data (Date and Sales)
            // -------------------------------------------------
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");

            DateTime startDate = new DateTime(2023, 1, 1);
            Random rnd = new Random();

            // Fill 10 rows of data
            for (int i = 0; i < 10; i++)
            {
                cells[i + 1, 0].PutValue(startDate.AddDays(i));   // Date column (A)
                cells[i + 1, 1].PutValue(rnd.Next(100, 500));    // Sales column (B)
            }

            // -------------------------------------------------
            // 2. Create a PivotTable that will serve as the Timeline source
            // -------------------------------------------------
            // Data range A1:B11, place pivot starting at D3
            int pivotIndex = sheet.PivotTables.Add("A1:B11", "D3", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Use Date as column field (required for Timeline) and Sales as data field
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // 3. Add a Timeline linked to the PivotTable
            // -------------------------------------------------
            // Place the Timeline at cell F1 (row 0, column 5)
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            // -------------------------------------------------
            // 4. Create a chart based on the original data
            // -------------------------------------------------
            // Add a column chart positioned from rows 15-30 and columns 0-10
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data series (Sales) and category (Date)
            chart.NSeries.Add("B2:B11", true);          // Sales values
            chart.NSeries.CategoryData = "A2:A11";     // Corresponding dates
            chart.Title.Text = "Sales Over Time";

            // -------------------------------------------------
            // 5. Export the chart to PDF
            // -------------------------------------------------
            string pdfPath = "SalesTimelineChart.pdf";
            chart.ToPdf(pdfPath);

            // -------------------------------------------------
            // 6. Save the workbook (optional, for verification)
            // -------------------------------------------------
            string xlsxPath = "SalesTimelineDemo.xlsx";
            workbook.Save(xlsxPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}