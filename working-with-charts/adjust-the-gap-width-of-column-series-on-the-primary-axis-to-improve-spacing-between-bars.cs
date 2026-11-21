// Title: Adjust column series gap width in a clustered column chart with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set gap width for clustered column chart | increase spacing between bars in Excel column chart using Aspose.Cells | how to change chart.GapWidth property in .NET workbook | customize column series spacing on primary axis with Aspose.Cells | adjust column chart bar gaps programmatically in C#
// Tags: adjust column chart gap width Aspose.Cells | set chart.GapWidth property .NET | clustered column chart spacing Excel | modify primary axis bar spacing Aspose.Cells | column series gap width C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// This example creates a workbook, adds sample data, inserts a clustered column chart, sets the chart's GapWidth property to 80 to widen the gaps between bars on the primary axis, and saves the file as AdjustedGapWidth.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a clustered column chart (rows: 5‑20, columns: 0‑10)
            // Use ChartType.Column for a standard clustered column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Adjust the gap width (0‑500). Lower values increase spacing between bars.
            chart.GapWidth = 80; // Example: set to 80 for wider gaps

            // Save the workbook (ensure the directory exists)
            string outputPath = "AdjustedGapWidth.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
