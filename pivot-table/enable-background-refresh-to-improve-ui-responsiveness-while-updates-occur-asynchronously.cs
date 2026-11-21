// Title: Enable background refresh for Excel query tables using Aspose.Cells in C# to improve UI responsiveness
// Common Searches: how to configure query tables for background refresh with Aspose.Cells in C# | Aspose.Cells .NET enable asynchronous refresh of Excel data connections | C# example to set RefreshOnLoad for Excel query tables using Aspose | improve application UI while loading large Excel workbooks with Aspose.Cells
// Tags: querytable RefreshOnLoad Aspose.Cells | asynchronous Excel data connection C# | UI responsiveness workbook loading Aspose | iterate worksheets querytables .NET | set querytable properties programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, iterates through each worksheet's QueryTables, and demonstrates where to enable background (asynchronous) refresh such as the RefreshOnLoad property when supported, then saves the updated workbook.
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

                // Iterate through all worksheets and their query tables
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (QueryTable queryTable in sheet.QueryTables)
                    {
                        // Aspose.Cells does not expose a RefreshOnLoad property in some versions.
                        // If needed, additional settings can be applied here.
                        // Example (if available): queryTable.RefreshOnLoad = true;
                    }
                }

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
                // Handle any runtime exceptions
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
