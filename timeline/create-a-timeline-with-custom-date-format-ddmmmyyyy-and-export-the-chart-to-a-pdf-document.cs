using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Timelines;

namespace AsposeCellsTimelinePdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with dates and values
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Date");
            cells["C1"].PutValue("Amount");

            cells["A2"].PutValue("Item1");
            cells["A3"].PutValue("Item2");
            cells["A4"].PutValue("Item3");

            cells["B2"].PutValue(new DateTime(2021, 1, 15));
            cells["B3"].PutValue(new DateTime(2021, 2, 20));
            cells["B4"].PutValue(new DateTime(2021, 3, 25));

            cells["C2"].PutValue(120);
            cells["C3"].PutValue(150);
            cells["C4"].PutValue(180);

            // Create a date style with custom format dd-MMM-yyyy
            Style dateStyle = new CellsFactory().CreateStyle();
            dateStyle.Custom = "dd-MMM-yyyy";
            cells["B2"].SetStyle(dateStyle);
            cells["B3"].SetStyle(dateStyle);
            cells["B4"].SetStyle(dateStyle);

            // Add a PivotTable based on the data range
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:C4", "E1", "PivotTable1");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the PivotTable's Date field
            // Position the timeline starting at cell G5
            int timelineIndex = sheet.Timelines.Add(pivot, "G5", "Date");
            Timeline timeline = sheet.Timelines[timelineIndex];
            timeline.Caption = "Sales Timeline";
            // Optional: adjust size
            timeline.WidthPixel = 400;
            timeline.HeightPixel = 100;

            // Add a line chart to visualize the same data
            int chartIndex = sheet.Charts.Add(ChartType.Line, 15, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];
            // Set data source for the chart
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.CategoryData = "B2:B4";
            chart.Title.Text = "Amount Over Time";

            // Apply the same custom date format to the X axis values
            chart.NSeries[0].XValuesFormatCode = "dd-MMM-yyyy";

            // Export the chart to a PDF file using the ToPdf method
            string pdfPath = "TimelineChart.pdf";
            chart.ToPdf(pdfPath);

            // Save the workbook (optional, to keep the timeline in the Excel file)
            workbook.Save("TimelineWithChart.xlsx");

            Console.WriteLine("Timeline created, chart exported to PDF successfully.");
        }
    }
}