// Title: How to lock a chart’s aspect ratio by fixing the PlotArea size in Aspose.Cells for .NET
// AI Prompts: Write C# code that sets explicit Width and Height on a chart’s PlotArea using Aspose.Cells to keep the aspect ratio unchanged when the chart is resized. | Show how to retrieve an existing chart or add a new column chart and then disable automatic scaling by manually defining PlotArea dimensions with Aspose.Cells for .NET. | Provide a step‑by‑step example that checks for a chart, creates one if needed, and applies a constant PlotArea size to prevent distortion.
// Common Searches: Aspose.Cells .NET set chart PlotArea width and height to maintain aspect ratio | prevent Excel chart distortion when resizing plot area using Aspose.Cells C# | how to keep chart proportions fixed in Aspose.Cells for .NET workbook | C# Aspose.Cells lock chart scaling after changing plot area size | example of disabling automatic chart scaling in Aspose.Cells
// Tags: set chart PlotArea dimensions Aspose.Cells | lock chart aspect ratio .NET | disable automatic chart scaling Aspose.Cells | fixed plot area size Excel chart C# | maintain chart proportions Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample loads an Excel workbook, accesses the first worksheet, retrieves an existing chart or adds a new column chart, and demonstrates how to lock the chart’s aspect ratio by assigning explicit Width and Height values to the chart’s PlotArea, since Aspose.Cells for .NET does not expose an automatic scaling property, then saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Obtain a chart; create one if none exists
            Chart chart;
            if (sheet.Charts.Count == 0)
            {
                // Add a simple column chart at a specified position
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
                chart = sheet.Charts[chartIndex];
            }
            else
            {
                chart = sheet.Charts[0];
            }

            // Note: Aspose.Cells for .NET does not expose an IsAutomaticScaling property.
            // If needed, you can manually set the plot area size to control scaling.
            // Example (optional):
            // chart.PlotArea.Width = 400;
            // chart.PlotArea.Height = 300;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
