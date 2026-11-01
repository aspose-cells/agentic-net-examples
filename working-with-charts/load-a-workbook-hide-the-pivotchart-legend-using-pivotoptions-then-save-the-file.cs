// Title: Hide the legend of a PivotChart in an Excel workbook using Aspose.Cells for .NET (C#) and save the file
// Common Searches: C# Aspose.Cells hide PivotChart legend before saving workbook | how to programmatically remove legend from an Excel PivotChart using Aspose.Cells | Aspose.Cells set ShowLegend false for PivotChart in .NET | example code to hide chart legend in a PivotChart with Aspose.Cells C# | save Excel file after modifying PivotChart legend using Aspose.Cells
// Tags: Aspose.Cells chart legend visibility | C# PivotChart legend manipulation | Excel workbook chart property update Aspose.Cells | Hide PivotChart legend programmatically .NET | Save modified workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, accesses the first chart (assumed to be a PivotChart) on the first worksheet, disables its legend by setting ShowLegend to false, ensures the output directory exists, and then saves the updated workbook to a new file.
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
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Assume the first chart in the first worksheet is a PivotChart
                Worksheet sheet = workbook.Worksheets[0];
                if (sheet.Charts.Count > 0)
                {
                    // Retrieve the first chart
                    Chart pivotChart = sheet.Charts[0];

                    // Hide the legend of the PivotChart using the correct API
                    pivotChart.ShowLegend = false;
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
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
