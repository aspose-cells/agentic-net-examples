// Title: Clearing all AutoFilter criteria from an Excel ListObject table using Aspose.Cells for .NET (C#)
// Common Searches: c# aspocells remove all filters from an Excel table programmatically | aspocells clear autfilter on listobject worksheet | reset excel table filters using Aspose.Cells .NET | how to refresh table after clearing filters with Aspose.Cells | remove table AutoFilter criteria before saving workbook in C#
// Tags: listobject autfilter reset aspocells | excel table filter removal c# aspocells | clear table criteria aspnet aspocells | refresh listobject after filter clear | save workbook without filters aspocells

using Aspose.Cells;
using Aspose.Cells.Tables;
using System;
using System.IO;

// The example loads an Excel workbook, accesses the first worksheet's first ListObject, clears every AutoFilter condition by emptying the FilterColumns collection, refreshes the filter to display all rows, and saves the updated workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the table with filters
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one table (ListObject)
            if (sheet.ListObjects.Count > 0)
            {
                // Get the first table in the worksheet
                ListObject table = sheet.ListObjects[0];

                // Clear all filters applied to the table, restoring full data visibility
                if (table.AutoFilter != null)
                {
                    // Remove all filter criteria
                    table.AutoFilter.FilterColumns.Clear();

                    // Refresh the filter to display all rows
                    table.AutoFilter.Refresh();
                }
                else
                {
                    Console.WriteLine("The table does not have an AutoFilter attached.");
                }
            }
            else
            {
                Console.WriteLine("No tables (ListObjects) found in the worksheet.");
            }

            // Save the workbook after clearing the filters
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
