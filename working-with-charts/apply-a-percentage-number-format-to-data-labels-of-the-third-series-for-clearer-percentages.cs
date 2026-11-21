// Title: Apply a percentage number format to the data labels of the third series in an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# format third chart series data labels as percentage | set percentage number format for chart series data labels in .NET | how to show percent values on specific series data labels with Aspose.Cells | C# code to apply custom number format to Excel chart series using Aspose.Cells | change data label format to 0.00% for third series in an Aspose.Cells chart
// Tags: Aspose.Cells chart series data label formatting | C# apply percentage number format to Excel chart labels | third series data label custom format Aspose.Cells | custom number format for chart series in .NET | 0.00% data label format Aspose.Cells | Excel chart data label percentage formatting C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads a workbook, accesses the first worksheet and its first chart, verifies a third series exists, enables data labels for that series, applies the "0.00%" number format to the labels, and saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "{InputFilePath}";
            string outputPath = "{OutputFilePath}";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("Error: No charts found in the first worksheet.");
                    return;
                }

                // Access the first chart on the worksheet (adjust index if needed)
                Chart chart = worksheet.Charts[0];

                // Ensure the chart has at least three series
                if (chart.NSeries.Count <= 2)
                {
                    Console.WriteLine("Error: The chart does not contain a third series.");
                    return;
                }

                // Access the third series (zero‑based index = 2)
                Series thirdSeries = chart.NSeries[2];

                // Enable data labels for the series
                thirdSeries.DataLabels.ShowValue = true;        // show values
                thirdSeries.DataLabels.ShowPercentage = true;   // show percentages (optional)

                // Apply percentage number format to the data labels
                thirdSeries.DataLabels.NumberFormat = "0.00%";

                // Ensure the output directory exists
                string? outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
