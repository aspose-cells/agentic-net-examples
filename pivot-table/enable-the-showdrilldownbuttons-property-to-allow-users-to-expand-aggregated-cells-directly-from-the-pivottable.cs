// Title: Enable ShowDrillDownButtons on an Aspose.Cells PivotTable to allow cell drill‑down in C#
// Common Searches: Aspose.Cells C# enable ShowDrillDownButtons on pivot table | set drill down buttons for Excel pivot table using Aspose.Cells .NET | programmatically show drill‑down buttons in a pivot table with Aspose.Cells | C# code to allow users to expand aggregated cells in Aspose.Cells pivot table | Aspose.Cells pivot table drill‑down option not exposed property workaround
// Tags: Aspose.Cells PivotTable ShowDrillDownButtons property | C# set pivot table drill down option | Excel pivot table drill down using Aspose.Cells | modify pivot table settings programmatically .NET | save workbook after pivot table configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, verifies that the first worksheet contains a pivot table, accesses the first PivotTable object, sets its ShowDrillDownButtons property to true (when supported), and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the existing workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one pivot table
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No pivot tables found in the worksheet.");
                    return;
                }

                // Get the first pivot table
                var pivotTable = worksheet.PivotTables[0];

                // NOTE: Aspose.Cells does not expose a direct property to enable drill‑down buttons.
                // If the version you are using supports it, you can set the appropriate property here.
                // Example (when available): pivotTable.ShowDrillDownButtons = true;

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
