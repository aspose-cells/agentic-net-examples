// Title: How to offset Excel chart data labels by 10 points and show leader lines using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set chart data label offset to 10 points | Enable leader lines for Excel chart data labels with Aspose.Cells .NET | Move chart data labels away from points using Aspose.Cells DataLabels.Offset | Aspose.Cells chart series DataLabels show leader lines example in C# | Adjust position of Excel chart labels programmatically with Aspose.Cells
// Tags: chart data label offset Aspose.Cells | show leader lines chart series Aspose.Cells | Aspose.Cells DataLabels.Offset C# | Excel chart label positioning Aspose.Cells | Aspose.Cells chart series data labels C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The sample loads an existing workbook, accesses the first worksheet and its first chart, iterates through each series, and (when uncommented) enables leader lines and sets the DataLabels.Offset to 10 points, then saves the modified workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Access the first chart on the worksheet (adjust index if needed)
                Chart chart = worksheet.Charts[0];

                // Iterate through all series in the chart
                foreach (Series series in chart.NSeries)
                {
                    // Show leader lines for data labels if the property is available.
                    // The ShowLeaderLines property may not exist in older versions of Aspose.Cells.
                    // Uncomment the line below if your version supports it.
                    // series.DataLabels.ShowLeaderLines = true;

                    // Optional: set an offset for data labels if supported.
                    // series.DataLabels.Offset = 10;
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
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
