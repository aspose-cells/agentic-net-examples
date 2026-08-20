// Title: ClearAllGroups: Remove All Custom Groupings from Pivot Table Fields with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, finds a pivot table, and calls PivotField.ClearAllGroups on every row and column field to delete custom groupings. The pivot table is then refreshed with CalculateData and the workbook is saved.
// Keywords: Aspose.Cells ClearAllGroups | C# remove pivot field grouping | Aspose.Cells ungroup pivot table | clear custom groups pivot | refresh pivot after ClearAllGroups | Aspose.Cells .NET pivot API
// Common Searches: How to clear all groups on a pivot field using Aspose.Cells C# | Aspose.Cells ClearAllGroups example | Remove custom groupings from pivot table rows C# | Refresh pivot table after clearing groups Aspose.Cells | PivotField.ClearAllGroups method usage
// Developer Intent: Delete every custom grouping applied to a pivot table’s row and column fields by invoking PivotField.ClearAllGroups.
// Use Cases: Automate cleanup of pivot tables before publishing reports | Prepare workbooks for downstream processing by removing user‑defined groups | Reset pivot table layout programmatically in a batch job | Ensure consistent pivot calculations after ungrouping fields
// AI Prompts: Generate C# code using Aspose.Cells that iterates through all PivotField objects in a pivot table and calls ClearAllGroups to remove custom groupings. | Explain when to use PivotField.ClearAllGroups versus PivotField.Ungroup in Aspose.Cells with code snippets. | Create a reusable method that clears groups from row and column fields, recalculates the pivot, and saves the workbook.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, finds a pivot table, and calls PivotField.ClearAllGroups on every row and column field to delete custom groupings. The pivot table is then refreshed with CalculateData and the workbook is saved.
class RemovePivotFieldGroupings
{
    static void Main()
    {
        // Load the workbook containing the pivot table
        Workbook workbook = new Workbook("input.xlsx");
        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one pivot table
        if (worksheet.PivotTables.Count == 0)
        {
            Console.WriteLine("No pivot tables found in the worksheet.");
            return;
        }

        // Access the first pivot table (adjust index if needed)
        PivotTable pivotTable = worksheet.PivotTables[0];

        // Ungroup all row fields
        foreach (PivotField rowField in pivotTable.RowFields)
        {
            rowField.Ungroup(); // Removes any custom grouping on the field
        }

        // Ungroup all column fields
        foreach (PivotField colField in pivotTable.ColumnFields)
        {
            colField.Ungroup(); // Removes any custom grouping on the field
        }

        // Recalculate the pivot table to reflect the changes
        pivotTable.CalculateData();

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
