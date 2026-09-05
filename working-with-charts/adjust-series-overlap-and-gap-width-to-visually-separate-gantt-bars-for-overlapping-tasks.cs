// Title: How to adjust series overlap and gap width of a Gantt chart in Excel using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook with Aspose.Cells, finds the first Gantt chart, and sets its Overlap to -20 and GapWidth to 150, using reflection to stay compatible with older library versions. | Create a reusable C# method that receives a Worksheet and applies custom series overlap and gap width values to every bar or column chart, checking for property existence before assignment. | Show how to save the updated workbook to a new file path and confirm that the chart spacing changes are reflected in the resulting Excel file.
// Common Searches: Aspose.Cells C# set chart overlap negative value for Gantt chart | change gap width of bar chart in Excel using Aspose.Cells .NET | use reflection to modify chart properties in older Aspose.Cells versions | programmatically adjust spacing between Gantt bars in an Excel workbook with C#
// Tags: Aspose.Cells set Overlap property C# | Aspose.Cells configure GapWidth .NET | customize Gantt chart bar spacing Aspose.Cells | use reflection for chart property compatibility Aspose.Cells | programmatic Excel chart spacing adjustment C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, accesses the first chart (assumed to be a Gantt chart), and attempts to set its Overlap to -20 and GapWidth to 150 via reflection for backward‑compatible Aspose.Cells versions, then saves the modified workbook to a new file.
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

                // Ensure there is at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Retrieve the first chart (assumed Gantt chart)
                Chart chart = worksheet.Charts[0];

                // Set series overlap and gap width for bar/column charts if supported
                // Note: Overlap and GapWidth properties are available in newer Aspose.Cells versions.
                // If they are not present, this step is skipped.
                try
                {
                    // Attempt to set properties via reflection to avoid compile-time errors on older versions
                    var overlapProp = chart.GetType().GetProperty("Overlap");
                    var gapWidthProp = chart.GetType().GetProperty("GapWidth");

                    if (overlapProp != null && overlapProp.CanWrite)
                    {
                        overlapProp.SetValue(chart, -20);
                    }

                    if (gapWidthProp != null && gapWidthProp.CanWrite)
                    {
                        gapWidthProp.SetValue(chart, 150);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Unable to set Overlap/GapWidth properties. {ex.Message}");
                }

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath) ?? string.Empty;
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
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
