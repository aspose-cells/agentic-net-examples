// Title: How to set a PivotTable’s default data source connection string using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set default connection string for PivotTable data source | Configure external data source for PivotTable using Aspose.Cells in .NET | Programmatically change PivotTable connection string with Aspose.Cells C# | Aspose.Cells PivotTable data source connection property missing workaround
// Tags: Aspose.Cells PivotTable data source configuration | C# modify PivotTable external connection | Aspose.Cells workbook PivotTable settings | default connection string handling Aspose.Cells | PivotTable data source API Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable class

namespace AsposeCellsExample
{
    // The example loads an existing workbook, checks for a PivotTable on the first worksheet, notes that the DataSourceConnectionString property is unavailable in the current Aspose.Cells version (so the connection string cannot be set directly), and then saves the workbook to a new file.
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
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure a PivotTable exists
                if (worksheet.PivotTables.Count == 0)
                {
                    Console.WriteLine("No PivotTable found in the first worksheet.");
                    return;
                }

                // Retrieve the first PivotTable
                PivotTable pivotTable = worksheet.PivotTables[0];

                // If needed, modify the data source connection string.
                // The DataSourceConnectionString property is not available in this version,
                // so this step is omitted or replaced with appropriate logic.

                // Ensure output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
