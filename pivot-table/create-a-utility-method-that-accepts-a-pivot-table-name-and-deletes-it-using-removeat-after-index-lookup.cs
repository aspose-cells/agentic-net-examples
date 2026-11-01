// Title: Delete a specific pivot table by name from an Excel workbook using Aspose.Cells for .NET with RemoveAt
// Common Searches: C# code to delete a pivot table by its name using Aspose.Cells | Aspose.Cells RemoveAt example for removing a named pivot table in .NET | How to programmatically delete a pivot table from an existing Excel file with Aspose.Cells | Find and remove a pivot table in a workbook using PivotTables.IndexOf in C# | Aspose.Cells .NET delete pivot table from workbook without opening Excel
// Tags: Aspose.Cells pivot table removal C# | RemoveAt PivotTables collection Aspose | Delete named pivot table Aspose.Cells | PivotTables.IndexOf usage Aspose.Cells | Excel workbook pivot table deletion .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The DeletePivotTable method loads the workbook, searches each worksheet for a pivot table matching the supplied name, obtains its index via PivotTables.IndexOf, removes it with PivotTables.RemoveAt, saves the workbook, and throws clear exceptions when the file or the specified pivot table cannot be found.
public static class PivotTableUtility
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="pivotTableName">Name of the pivot table to delete.</param>
    public static void DeletePivotTable(string filePath, string pivotTableName)
    {
        // Ensure the file exists before attempting to load it
        if (!File.Exists(filePath))
            throw new FileNotFoundException($"File not found: {filePath}");

        try
        {
            // Load the workbook (lifecycle load rule)
            Workbook workbook = new Workbook(filePath);

            // Flag to indicate if the pivot table was found
            bool found = false;

            // Iterate through all worksheets to locate the pivot table
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Retrieve the pivot table by name
                PivotTable pt = sheet.PivotTables[pivotTableName];
                if (pt != null)
                {
                    // Get the index of the pivot table within the collection
                    int index = sheet.PivotTables.IndexOf(pt);

                    // Remove the pivot table using RemoveAt (as requested)
                    sheet.PivotTables.RemoveAt(index);

                    found = true;
                    break; // Pivot table names are unique within a workbook, so we can exit
                }
            }

            if (!found)
                throw new ArgumentException($"Pivot table '{pivotTableName}' not found in the workbook.");

            // Save the modified workbook (lifecycle save rule)
            workbook.Save(filePath);
        }
        catch (Exception ex)
        {
            // Log and rethrow for the caller to handle
            Console.Error.WriteLine($"Error deleting pivot table: {ex.Message}");
            throw;
        }
    }

    // Entry point for console execution
    public static void Main(string[] args)
    {
        // Expect two arguments: the Excel file path and the pivot table name
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: PivotTableUtility <excelFilePath> <pivotTableName>");
            return;
        }

        string filePath = args[0];
        string pivotName = args[1];

        try
        {
            DeletePivotTable(filePath, pivotName);
            Console.WriteLine("Pivot table deleted successfully.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Operation failed: {ex.Message}");
        }
    }
}
