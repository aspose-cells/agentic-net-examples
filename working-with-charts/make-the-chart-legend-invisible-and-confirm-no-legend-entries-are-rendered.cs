// Title: Hide a chart legend and verify it is not rendered using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide chart legend and check visibility | how to programmatically remove legend from Excel chart with Aspose.Cells | verify legend is not displayed after setting ShowLegend false in Aspose.Cells | C# Aspose.Cells hide chart legend without affecting data series | Aspose.Cells chart legend invisible validation example
// Tags: Aspose.Cells chart legend visibility control | ShowLegend property usage Aspose.Cells | programmatic removal of Excel chart legend | C# validation of hidden chart legend | Aspose.Cells workbook chart manipulation

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example loads an existing Excel workbook, ensures a chart exists on the first worksheet, sets the chart's ShowLegend property to false to hide the legend, confirms the legend is invisible, creates the output directory if needed, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
                return;
            }

            // Get the first chart on the worksheet
            Chart chart = sheet.Charts[0];

            // Hide the chart legend
            chart.ShowLegend = false;

            // Confirm that the legend is hidden
            if (!chart.ShowLegend)
            {
                Console.WriteLine("Legend is invisible as expected.");
            }
            else
            {
                Console.WriteLine("Legend is still visible.");
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
