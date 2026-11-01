// Title: How to make chart legend background transparent while preserving text color using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# remove chart legend background fill without changing label color | set legend area transparent in an Excel chart with Aspose.Cells .NET | clear legend background color programmatically in a workbook chart using Aspose.Cells | make chart legend background invisible while keeping text color in C# | transparent legend background for Excel chart via Aspose.Cells API
// Tags: chart legend background transparency Aspose.Cells | legend area fill removal .NET | Aspose.Cells chart legend styling | transparent legend background C# | Excel chart legend formatting Aspose.Cells

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing; // For Area related properties

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses the first chart on the first worksheet, and sets both the foreground and background colors of the chart's legend area to Transparent. This removes the legend's background fill while leaving the text color unchanged, then saves the modified workbook.
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

                // Assume the chart is on the first worksheet and is the first chart object
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found on the first worksheet.");
                    return;
                }

                Chart chart = sheet.Charts[0];

                // Remove background fill from the legend (applies to all legend entries)
                // Set both foreground and background colors to Transparent
                chart.Legend.Area.ForegroundColor = Color.Transparent;
                chart.Legend.Area.BackgroundColor = Color.Transparent;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
