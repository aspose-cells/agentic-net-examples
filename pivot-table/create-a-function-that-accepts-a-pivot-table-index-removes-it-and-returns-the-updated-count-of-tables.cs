// Title: Remove a specific pivot table by index and get the remaining count with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete pivot table at given index and return count | how to remove a pivot table from the first worksheet using Aspose.Cells | Aspose.Cells remove pivot table by zero based index and get remaining tables | count remaining pivot tables after removal in Excel using Aspose.Cells C# | remove pivot table programmatically with Aspose.Cells and save workbook
// Tags: Aspose.Cells remove pivot table by index | C# pivot table deletion Aspose.Cells | update workbook after pivot table removal | retrieve remaining pivot table count Aspose.Cells | Excel worksheet pivot table management C#

using System;
using System.IO;
using Aspose.Cells;

// The RemovePivotTable method loads an Excel workbook, accesses the first worksheet, validates the zero‑based pivot table index, removes the specified pivot table, saves the workbook, and returns the updated number of pivot tables in that worksheet, or -1 on error.
public static class PivotTableHelper
{
    /// <param name="filePath">Full path to the Excel file.</param>
    /// <param name="pivotTableIndex">Zero‑based index of the pivot table to remove.</param>
    /// <returns>
    /// Number of remaining pivot tables in the first worksheet after removal,
    /// or -1 if an error occurs.
    /// </returns>
    public static int RemovePivotTable(string filePath, int pivotTableIndex)
    {
        try
        {
            // Load the workbook from the given file path.
            Workbook workbook = new Workbook(filePath);

            // Access the first worksheet (adjust as needed for other sheets).
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the index is within the valid range.
            if (pivotTableIndex >= 0 && pivotTableIndex < sheet.PivotTables.Count)
            {
                // Remove the pivot table at the specified index.
                sheet.PivotTables.RemoveAt(pivotTableIndex);
            }

            // Save the workbook to persist the changes.
            workbook.Save(filePath);

            // Return the updated count of pivot tables in the worksheet.
            return sheet.PivotTables.Count;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error removing pivot table: {ex.Message}");
            return -1;
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        // Expecting two arguments: file path and pivot table index.
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: <exe> <ExcelFilePath> <PivotTableIndex>");
            return;
        }

        string filePath = args[0];
        if (!File.Exists(filePath))
        {
            Console.Error.WriteLine($"File not found: {filePath}");
            return;
        }

        if (!int.TryParse(args[1], out int pivotIndex))
        {
            Console.Error.WriteLine($"Invalid pivot table index: {args[1]}");
            return;
        }

        try
        {
            int remaining = PivotTableHelper.RemovePivotTable(filePath, pivotIndex);
            if (remaining >= 0)
            {
                Console.WriteLine($"Pivot table removed. Remaining pivot tables: {remaining}");
            }
            else
            {
                Console.Error.WriteLine("Failed to remove pivot table.");
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
