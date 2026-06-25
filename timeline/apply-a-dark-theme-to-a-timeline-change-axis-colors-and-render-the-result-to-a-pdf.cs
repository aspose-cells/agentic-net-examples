using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Pivot;
using Aspose.Cells.Timelines;

class TimelineDarkThemePdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate worksheet with date and value data
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue(new DateTime(2023, 2, 1));
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue(new DateTime(2023, 3, 1));
            sheet.Cells["B4"].PutValue(200);

            // Create a pivot table based on the data
            PivotTableCollection pivots = sheet.PivotTables;
            int pivotIndex = pivots.Add("A1:B4", "D1", "Pivot1");
            PivotTable pivot = pivots[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Value");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a Timeline linked to the pivot table
            sheet.Timelines.Add(pivot, 10, 5, "Date");

            // Add a chart to visualize the same data
            int chartIndex = sheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);               // Values
            chart.NSeries.CategoryData = "A2:A4";           // Categories (dates)

            // Save the workbook as a PDF with a dark theme applied to the timeline
            string outputPath = "TimelineDarkTheme.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}