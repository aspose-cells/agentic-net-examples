// Title: Determine whether an Excel chart has a primary value axis using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel workbook, retrieves the first chart, and uses chart.HasAxis to report if a primary value axis is present. | Show how to extend the example to also check for a secondary category axis on the same chart and output both results. | Create a resilient C# snippet that iterates over all charts in a worksheet, uses chart.HasAxis to detect value axes, and logs the findings while handling missing files and empty chart collections.
// Common Searches: asp.net aspose.cells check primary value axis in chart c# | how to use chart.HasAxis to detect value axis in Excel file with Aspose.Cells | c# code sample for verifying existence of a value axis on an Excel chart | Aspose.Cells chart.HasAxis AxisType.Value true example | detect missing value axis in Excel chart using Aspose.Cells for .NET
// Tags: Aspose.Cells chart.HasAxis value axis detection | C# verify primary value axis in Excel chart | Aspose.Cells check chart axis existence | Excel chart value axis verification using Aspose.Cells | AxisType.Value primary axis check Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, accesses the first worksheet and its first chart, then calls chart.HasAxis(AxisType.Value, true) to determine if a primary value axis exists, outputting the boolean result.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            try
            {
                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Retrieve the first chart
                Chart chart = worksheet.Charts[0];

                // Determine whether the chart contains a primary value axis
                bool hasValueAxis = chart.HasAxis(AxisType.Value, true);

                // Output the result
                Console.WriteLine($"Chart has value axis: {hasValueAxis}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
