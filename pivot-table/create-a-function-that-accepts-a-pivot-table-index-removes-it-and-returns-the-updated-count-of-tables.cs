// Title: Aspose.Cells for .NET – Remove a Pivot Table by Index and Get Updated Count
// Description: C# example that creates a workbook, adds sample pivot tables, validates an index, removes the pivot table at that position using PivotTables.RemoveAt, and returns the remaining pivot table count.
// Keywords: Aspose.Cells remove pivot table | PivotTables.RemoveAt C# | delete pivot table by index | pivot table count after removal | Aspose.Cells .NET pivot management
// Common Searches: how to delete a pivot table at a specific index using Aspose.Cells | Aspose.Cells get number of pivot tables after removal | C# remove second pivot table from worksheet | PivotTables.RemoveAt usage example
// Developer Intent: Programmatically delete a pivot table at a given zero‑based index from the first worksheet and obtain the new total of pivot tables.
// Use Cases: Clean up automatically generated pivot tables before exporting a workbook. | Validate user‑driven deletions in a reporting dashboard and adjust UI accordingly. | Reduce file size by removing unwanted pivot tables during batch processing.
// AI Prompts: Generate a C# method that takes an integer index, checks bounds, removes the corresponding pivot table with Aspose.Cells, and returns the updated count. | Write unit tests for RemovePivotTableAtIndex covering valid indices, out‑of‑range errors, and the case where all pivot tables are removed. | Explain the exceptions thrown by PivotTables.RemoveAt in Aspose.Cells and best practices for handling them.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, adds sample pivot tables, validates an index, removes the pivot table at that position using PivotTables.RemoveAt, and returns the remaining pivot table count.
public class PivotTableHelper
{
    // Removes the pivot table at the specified index from the first worksheet
    // and returns the updated count of pivot tables in that worksheet.
    public static int RemovePivotTableAtIndex(int indexToRemove)
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for pivot tables
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(200);
        sheet.Cells["B4"].PutValue(300);

        // Add three pivot tables to demonstrate removal
        sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        sheet.PivotTables.Add("A1:B4", "D10", "PivotTable2");
        sheet.PivotTables.Add("A1:B4", "D20", "PivotTable3");

        // Ensure the index is within range to avoid ArgumentOutOfRangeException
        if (indexToRemove < 0 || indexToRemove >= sheet.PivotTables.Count)
        {
            throw new ArgumentOutOfRangeException(nameof(indexToRemove),
                $"Index must be between 0 and {sheet.PivotTables.Count - 1}");
        }

        // Remove the pivot table at the specified index using the documented RemoveAt method
        sheet.PivotTables.RemoveAt(indexToRemove);

        // Return the updated count of pivot tables
        return sheet.PivotTables.Count;
    }

    // Example usage
    public static void Main()
    {
        int updatedCount = RemovePivotTableAtIndex(1); // Remove the second pivot table
        Console.WriteLine("Updated pivot tables count: " + updatedCount);
    }
}
