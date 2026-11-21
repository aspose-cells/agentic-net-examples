// Title: C# example: Using Aspose.Cells chart.HasAxis to verify a primary category (X) axis in an Excel chart
// Common Searches: C# Aspose.Cells how to determine if an Excel chart contains a primary category axis | using chart.HasAxis with AxisType.Category to check X axis presence in Aspose.Cells | sample code to detect category axis in a workbook chart with Aspose.Cells for .NET | verify existence of primary X axis in a chart programmatically using Aspose.Cells | Aspose.Cells chart.HasAxis method example for category axis detection in C#
// Tags: Aspose.Cells chart.HasAxis Category axis C# | detect primary X axis Aspose.Cells | Excel chart axis existence check .NET | AxisType.Category usage Aspose.Cells | C# chart axis verification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The C# program loads an Excel workbook, accesses the first worksheet, ensures a chart exists, retrieves the first chart, and calls chart.HasAxis(AxisType.Category, true) to determine whether a primary category (X) axis is present, outputting the result before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                }
                else
                {
                    // Access the first chart
                    Chart chart = worksheet.Charts[0];

                    // Check if a primary category (X) axis exists
                    bool hasCategoryAxis = chart.HasAxis(AxisType.Category, true);

                    Console.WriteLine($"Chart has category axis: {hasCategoryAxis}");
                }

                // Save the workbook (even if unchanged)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
