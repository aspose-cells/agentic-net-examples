// Title: Place an Excel chart legend at the bottom‑right corner and apply Calibri 10 font with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart legend to bottom right corner | change legend font to Calibri size 10 in Aspose.Cells chart | use XRatioToChart to move Excel chart legend in .NET | disable automatic legend size and define custom dimensions Aspose.Cells | position Excel chart legend at bottom‑right using Aspose.Cells API
// Tags: chart legend positioning Aspose.Cells | legend font customization Calibri Aspose.Cells | bottom right legend placement .NET Excel | XRatioToChart YRatioToChart usage Aspose.Cells | custom legend size Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartLegendPositioning
{
    // The example creates a workbook, adds a column chart, and configures the chart legend to appear in the bottom‑right corner by setting Position to Bottom, disabling automatic sizing, specifying Width and Height, and adjusting XRatioToChart/YRatioToChart. It also changes the legend font to Calibri size 10 before saving the file as an XLSX workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the legend: place it at the bottom‑right corner
                Legend legend = chart.Legend;
                legend.Position = LegendPositionType.Bottom;
                legend.IsAutomaticSize = false;
                legend.Width = 300;   // Width in points
                legend.Height = 100;  // Height in points

                // Use ratio properties to move legend toward the right side
                legend.XRatioToChart = 0.7; // Approx. 70% from the left edge
                legend.YRatioToChart = 0.9; // Near the bottom edge

                // Set legend font to Calibri, size 10
                legend.Font.Name = "Calibri";
                legend.Font.Size = 10;

                // Save the workbook
                string outputPath = "ChartWithLegendBottomRight.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved to '{outputPath}'. Legend positioned at bottom‑right with Calibri 10 font.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
