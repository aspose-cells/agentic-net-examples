// Title: Clear all AutoFilter criteria from a ListObject (Excel table) using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, retrieves the first ListObject, removes every AutoFilter condition, refreshes the worksheet filter, and saves the workbook with Aspose.Cells. | Show a .NET example that programmatically resets all filter columns of an Excel table (ListObject) and persists the changes using Aspose.Cells.
// Common Searches: Aspose.Cells C# clear filters on Excel ListObject | remove all AutoFilter criteria from a worksheet table using Aspose.Cells .NET | reset Excel table filters programmatically with Aspose.Cells | how to refresh worksheet autofilter after clearing filters in C#
// Tags: clear ListObject autofilter Aspose.Cells | reset Excel table filters C# | Aspose.Cells refresh worksheet autofilter | save workbook after filter removal .NET | programmatic table filter management Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsExample
{
    // The sample loads an Excel workbook, accesses the first ListObject on the first worksheet, clears any AutoFilter criteria by emptying the FilterColumns collection, refreshes the worksheet AutoFilter, and saves the modified file.
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
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook containing the table with filters
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index or name as needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first table (ListObject) on the worksheet
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables (ListObjects) found on the worksheet.");
                    return;
                }

                ListObject table = worksheet.ListObjects[0];

                // Clear any filter criteria applied to the table
                if (table.AutoFilter != null)
                {
                    // Remove all filter columns
                    table.AutoFilter.FilterColumns.Clear();

                    // Refresh the worksheet's autofilter to display all rows
                    worksheet.AutoFilter.Refresh();
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook; the table will now display all rows without any filters
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
