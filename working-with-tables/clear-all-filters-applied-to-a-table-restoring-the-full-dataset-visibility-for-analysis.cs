// Title: Clear all filters from an Excel table (ListObject) using Aspose.Cells for .NET
// Common Searches: asp.net clear excel table filters using Aspose.Cells | c# remove all auto filter criteria from a ListObject in a workbook | programmatically reset excel table filters with Aspose.Cells .NET | how to refresh worksheet after clearing table filters in Aspose.Cells | example code to clear filters on an Excel ListObject using Aspose.Cells
// Tags: Aspose.Cells reset ListObject auto filter | C# clear Excel table filter criteria Aspose.Cells | Aspose.Cells refresh worksheet after filter removal | Excel ListObject filter reset using Aspose.Cells .NET | Aspose.Cells save workbook after clearing filters

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads a workbook, checks for a ListObject on the first worksheet, accesses its AutoFilter, clears all filter columns, refreshes the filter to display all rows, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "Input.xlsx";
            string outputPath = "Output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"Input file not found: {inputPath}");
                }

                // Load the workbook
                var workbook = new Workbook(inputPath);

                // Access the first worksheet
                var worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one table (ListObject)
                if (worksheet.ListObjects.Count == 0)
                {
                    Console.WriteLine("No tables found in the worksheet.");
                }
                else
                {
                    // Get the first table
                    var table = worksheet.ListObjects[0];

                    // Clear all filters applied to the table
                    var autoFilter = table.AutoFilter;
                    if (autoFilter != null)
                    {
                        // Remove any filter criteria
                        autoFilter.FilterColumns.Clear();
                        // Refresh to show all rows
                        autoFilter.Refresh();
                    }
                }

                // Save the workbook with filters cleared
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
