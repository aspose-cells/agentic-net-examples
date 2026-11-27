// Title: How to add a vertical white‑to‑gray gradient to a chart’s plot area with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set vertical gradient fill on chart plot area | C# code to apply two‑color gradient to chart background using Aspose.Cells | Create chart with white to gray gradient plot area in Aspose.Cells .NET | Set plot area FillFormat gradient style vertical Aspose.Cells example | Aspose.Cells chart plot area gradient color transition white gray
// Tags: vertical two‑color gradient fill Aspose.Cells | chart plot area fill format C# | Aspose.Cells gradient style vertical | apply gradient to chart area .NET | SetTwoColorGradient Aspose.Cells example

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// This example creates a workbook, adds a column chart with sample data, accesses the chart's PlotArea, and uses the Area.FillFormat.SetTwoColorGradient method to apply a vertical gradient that transitions from white at the top to gray at the bottom. It also demonstrates setting a black border for the plot area and saving the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the plot area of the chart
            PlotArea plotArea = chart.PlotArea;

            // Apply a vertical gradient fill to the plot area:
            // White at the top, Gray at the bottom
            Area plotAreaArea = plotArea.Area;
            plotAreaArea.FillFormat.SetTwoColorGradient(
                Color.White,          // Start color (top)
                Color.Gray,           // End color (bottom)
                GradientStyleType.Vertical, // Gradient direction
                1);                   // Variant (1 is the default)

            // Optional: make the plot area border visible
            plotArea.Border.Color = Color.Black;

            // Save the workbook
            string outputPath = "ChartWithGradientPlotArea.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
