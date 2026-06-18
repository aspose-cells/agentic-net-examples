using System;
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
            // 1. Populate sample data (Date, Category, Value)
            // -------------------------------------------------
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Category");
            cells["C1"].PutValue("Value");

            DateTime startDate = new DateTime(2023, 1, 1);
            string[] categories = { "A", "B", "C" };
            Random rnd = new Random();

            for (int i = 0; i < 10; i++)
            {
                cells[i + 1, 0].PutValue(startDate.AddDays(i));                     // Date column
                cells[i + 1, 1].PutValue(categories[i % categories.Length]);      // Category column
                cells[i + 1, 2].PutValue(rnd.Next(50, 200));                       // Value column
            }

            // -------------------------------------------------
            // 2. Create a PivotTable based on the data range
            // -------------------------------------------------
            int pivotIndex = sheet.PivotTables.Add("A1:C11", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIndex];

            // Add fields to appropriate areas
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Column, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");

            // Timeline requires the date field to be in the Page (filter) area
            pivot.AddFieldToArea(PivotFieldType.Page, "Date");

            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // 3. Add a Timeline linked to the Date field of the PivotTable
            // -------------------------------------------------
            int timelineIndex = sheet.Timelines.Add(pivot, 0, 5, "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";

            // -------------------------------------------------
            // 4. Create a chart that visualizes the same data
            // -------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("C2:C11", true);          // Values
            chart.NSeries.CategoryData = "A2:A11";      // Dates as categories
            chart.Title.Text = "Sales by Date and Category";

            // -------------------------------------------------
            // 5. Export the chart to a PDF file
            // -------------------------------------------------
            chart.ToPdf("SalesTimelineChart.pdf");

            // Save the workbook to verify the timeline and chart
            workbook.Save("SalesTimelineWorkbook.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}