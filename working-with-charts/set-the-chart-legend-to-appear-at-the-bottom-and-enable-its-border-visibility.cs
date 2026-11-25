// Title: How to position a chart legend at the bottom and display its border using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend to bottom | enable legend border visibility in an Aspose.Cells chart | Aspose.Cells set legend border color and weight in .NET | move chart legend to bottom and auto‑size it with Aspose.Cells | C# Aspose.Cells show legend border with WideLine weight
// Tags: chart legend position bottom Aspose.Cells | legend border visibility Aspose.Cells C# | set legend border weight Aspose.Cells | automatic legend sizing Aspose.Cells | column chart legend formatting Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace ChartLegendBottomBorderDemo
{
    // The example creates a workbook, adds a column chart, moves the legend to the bottom, makes the legend border visible with a black color and WideLine weight, enables automatic legend sizing, and saves the file as an XLSX using Aspose.Cells for .NET.
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
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Access the legend and configure it
                Legend legend = chart.Legend;

                // Position the legend at the bottom of the chart
                legend.Position = LegendPositionType.Bottom;

                // Enable the legend border visibility
                legend.Border.IsVisible = true;
                legend.Border.Color = System.Drawing.Color.Black;
                // Use a predefined weight type (e.g., WideLine) – Weight property expects WeightType enum
                legend.Border.Weight = WeightType.WideLine;

                // Optional: make the legend size automatic so it fits the bottom area
                legend.IsAutomaticSize = true;

                // Ensure the output directory exists
                string outputPath = "ChartLegendBottomBorder.xlsx";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                try
                {
                    workbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine("Error saving the workbook:");
                    Console.WriteLine(saveEx.Message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the chart:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
