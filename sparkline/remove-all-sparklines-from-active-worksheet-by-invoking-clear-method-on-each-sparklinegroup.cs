// Title: How to remove all sparkline groups from the active worksheet using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells clear all sparkline groups in the active sheet | How to delete sparkline groups from an Excel workbook with Aspose.Cells | Remove sparklines from a worksheet programmatically using Aspose.Cells .NET | Aspose.Cells SparklineGroups.Clear example in C#
// Tags: Aspose.Cells SparklineGroups.Clear method | C# remove Excel sparkline groups | clear sparkline groups active worksheet | Aspose.Cells delete sparklines from workbook | Excel sparkline removal using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file, accesses the active worksheet, clears every SparklineGroup on that sheet with the SparklineGroups.Clear() method, ensures the output directory exists, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Get the active worksheet
                Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

                // Remove all sparkline groups from the active sheet
                // SparklineGroup type may not be available in older versions; use collection Clear method
                activeSheet.SparklineGroups.Clear();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
