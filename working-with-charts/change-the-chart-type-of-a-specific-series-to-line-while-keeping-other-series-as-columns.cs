// Title: Change a single chart series to a line type while keeping other series as columns with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change only one chart series to line type | set specific series to line chart in an existing Excel workbook using Aspose.Cells | keep other chart series as column while converting one series to line in .NET | modify chart series by index without recreating the chart Aspose.Cells | change chart series type without affecting other series in C#
// Tags: Aspose.Cells modify chart series type | C# change series to line chart | preserve other series as column chart | chart series index manipulation Aspose.Cells | update existing chart series .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, checks for a worksheet and chart, then changes the chart series at index 1 to a line type while leaving all other series as columns, and finally saves the modified workbook.
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
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure the workbook contains at least one worksheet
                if (workbook.Worksheets.Count == 0)
                {
                    Console.WriteLine("The workbook does not contain any worksheets.");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("The worksheet does not contain any charts.");
                    return;
                }

                // Retrieve the first chart on the worksheet
                Chart chart = worksheet.Charts[0];

                // Index of the series to modify (0‑based)
                int seriesIndex = 1;

                // Verify the series index is within range
                if (seriesIndex >= 0 && seriesIndex < chart.NSeries.Count)
                {
                    // Change the series type to Line while other series remain unchanged
                    chart.NSeries[seriesIndex].Type = ChartType.Line;
                }
                else
                {
                    Console.WriteLine($"Series index {seriesIndex} is out of range. The chart contains {chart.NSeries.Count} series.");
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
