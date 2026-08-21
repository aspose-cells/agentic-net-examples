// Title: C# – Delete a Pivot Table by Name Using Aspose.Cells RemoveAt
// Description: A concise C# helper that validates inputs, finds a pivot table by its Name property, retrieves its index, and removes it from the worksheet’s PivotTableCollection with RemoveAt. Includes a demo that loads a workbook, deletes "PivotTable1" from the first sheet, and saves the result.
// Keywords: Aspose.Cells delete pivot table C# | remove pivot table by name .NET | PivotTableCollection RemoveAt | C# find pivot table index | Aspose.Cells PivotTable helper | delete specific pivot table worksheet | Aspose.Cells example GitHub | automated workbook cleanup
// Common Searches: how to delete a pivot table by name in Aspose.Cells | remove pivot table using RemoveAt Aspose.Cells .NET | find pivot table index and delete it C# | Aspose.Cells delete specific pivot table example | C# code to remove pivot table from worksheet
// Developer Intent: Remove a specific pivot table identified by its name from an Excel worksheet.
// Use Cases: Erase an obsolete pivot table before recreating it with refreshed data. | Clean up generated reports by programmatically deleting unused pivot tables. | Run a batch job that scans multiple sheets and removes pivot tables matching a naming convention.
// AI Prompts: Write a C# method that deletes a pivot table by name using Aspose.Cells and throws a custom exception when the table is missing. | Show how to modify DeletePivotTableByName to log to a file instead of the console. | Provide code that iterates through all worksheets in a workbook and calls DeletePivotTableByName for every pivot table named "SalesPivot".

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsUtilities
{
    // A concise C# helper that validates inputs, finds a pivot table by its Name property, retrieves its index, and removes it from the worksheet’s PivotTableCollection with RemoveAt. Includes a demo that loads a workbook, deletes "PivotTable1" from the first sheet, and saves the result.
    public static class PivotTableHelper
    {
        /// <param name="worksheet">The worksheet containing the pivot table.</param>
        /// <param name="pivotTableName">The name of the pivot table to delete.</param>
        public static void DeletePivotTableByName(Worksheet worksheet, string pivotTableName)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (string.IsNullOrEmpty(pivotTableName)) throw new ArgumentException("Pivot table name cannot be null or empty.", nameof(pivotTableName));

            PivotTableCollection pivots = worksheet.PivotTables;
            int targetIndex = -1;

            // Locate the pivot table index by matching its Name property
            for (int i = 0; i < pivots.Count; i++)
            {
                if (pivots[i].Name.Equals(pivotTableName, StringComparison.OrdinalIgnoreCase))
                {
                    targetIndex = i;
                    break;
                }
            }

            // If found, remove it using RemoveAt
            if (targetIndex >= 0)
            {
                pivots.RemoveAt(targetIndex);
            }
            else
            {
                // Optionally handle the case where the pivot table does not exist
                Console.WriteLine($"Pivot table \"{pivotTableName}\" not found in worksheet \"{worksheet.Name}\".");
            }
        }

        // Example usage
        public static void Demo()
        {
            const string inputPath = "InputWorkbook.xlsx";
            const string outputPath = "OutputWorkbook.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load an existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the pivot table resides in the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Delete the pivot table named "PivotTable1"
            DeletePivotTableByName(sheet, "PivotTable1");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
    }

    // Entry point for the application
    public static class Program
    {
        public static void Main()
        {
            try
            {
                PivotTableHelper.Demo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
