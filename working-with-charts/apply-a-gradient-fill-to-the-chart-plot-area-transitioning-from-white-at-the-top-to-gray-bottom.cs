// Title: How to apply a vertical white‑to‑gray two‑color gradient fill to a chart plot area using Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add sample data, insert a column chart, and use FillFormat.SetTwoColorGradient to set a vertical white‑to‑gray gradient on the chart's plot area in C#. | Show how to configure a chart's PlotArea.FillFormat with a vertical two‑color gradient (white at the top, gray at the bottom) using the Aspose.Cells .NET API. | Write C# code that builds an Excel file, adds a column chart, applies a vertical gradient background to the plot area, and saves the workbook.
// Common Searches: Aspose.Cells C# set vertical gradient on chart plot area | two‑color gradient fill for Excel chart using Aspose.Cells .NET | how to make chart background fade from white to gray with Aspose.Cells
// Tags: Aspose.Cells chart plot area gradient | SetTwoColorGradient C# | vertical two‑color gradient fill | column chart background gradient Aspose.Cells | FillFormat gradient style vertical

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

// The example creates a workbook, populates it with data, adds a column chart, and applies a vertical two‑color gradient (white at the top, gray at the bottom) to the chart's plot area using Aspose.Cells' FillFormat.SetTwoColorGradient method, then saves the file as an Excel workbook.
class GradientChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Access the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Add some sample data for the chart.
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart data source.
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Apply a vertical two‑color gradient fill to the plot area:
            // White at the top, gray at the bottom.
            chart.PlotArea.Area.FillFormat.SetTwoColorGradient(
                Color.White,          // Start color (top)
                Color.Gray,           // End color (bottom)
                GradientStyleType.Vertical,
                1);                    // Variant (default)

            // Optional: set a title for clarity.
            chart.Title.Text = "Gradient Plot Area Example";

            // Save the workbook to a file.
            string outputPath = "GradientChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
