using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Timelines;
using Aspose.Cells.Drawing;

class UpdateTimelineAndExportPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: Date, Sales, Profit
            cells["A1"].PutValue("Date");
            cells["B1"].PutValue("Sales");
            cells["C1"].PutValue("Profit");
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 5; i++)
            {
                cells[i + 1, 0].PutValue(startDate.AddMonths(i));
                cells[i + 1, 1].PutValue(1000 + i * 200);   // Sales
                cells[i + 1, 2].PutValue(200 + i * 50);     // Profit
            }

            // Create a pivot table using the data range
            int pivotIdx = sheet.PivotTables.Add("A1:C6", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Place the Date field in the column area (required for Timeline)
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            // Add Sales and Profit as data fields
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.AddFieldToArea(PivotFieldType.Data, "Profit");

            // Refresh and calculate the pivot table
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a timeline linked to the Date field of the pivot table
            int timelineIdx = sheet.Timelines.Add(pivot, "G1", "Date");
            Timeline timeline = sheet.Timelines[timelineIdx];
            timeline.Caption = "Sales & Profit Timeline";
            timeline.Name = "SalesProfitTimeline";

            // Create a column chart based on the pivot table data
            int chartIdx = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart chart = sheet.Charts[chartIdx];

            // First series: Sales
            chart.NSeries.Add("=PivotTable1!Sales", true);
            // Second series: Profit
            chart.NSeries.Add("=PivotTable1!Profit", true);
            // Category (X‑axis) data: Date
            chart.NSeries.CategoryData = "=PivotTable1!Date";
            chart.Title.Text = "Sales and Profit Over Time";

            // Add a textbox annotation to the worksheet
            Shape annotation = sheet.Shapes.AddTextBox(5, 0, 5, 0, 200, 50);
            annotation.Text = "Generated on " + DateTime.Now.ToString("yyyy-MM-dd");
            annotation.Font.Size = 10;
            annotation.Font.Color = System.Drawing.Color.Blue;

            // Export the chart (which includes the timeline) to PDF
            chart.ToPdf("TimelineChart.pdf");

            // Save the workbook (optional, for verification)
            workbook.Save("TimelineWithChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}