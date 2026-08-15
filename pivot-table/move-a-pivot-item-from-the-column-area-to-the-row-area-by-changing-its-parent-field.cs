// Title: Aspose.Cells for .NET – Move a Pivot Table Field from Column to Row
// Description: C# example that loads an Excel workbook, locates the first pivot table, removes a specified column field, adds the same field to the row area, refreshes the pivot cache, recalculates data, and saves the modified file using Aspose.Cells.
// Keywords: Aspose.Cells pivot field relocation | C# move pivot column to row | programmatic pivot layout change .NET | remove column field Aspose.Cells | add row field Aspose.Cells | refresh pivot table after layout edit
// Common Searches: how to change pivot table field orientation with Aspose.Cells | C# code to move pivot column field to rows | Aspose.Cells move pivot field programmatically | update pivot layout dynamically in .NET | convert column field to row field in Excel using Aspose
// Developer Intent: Reassign a pivot table field from the Column area to the Row area via code.
// Use Cases: Switch a dimension (e.g., Region) from horizontal to vertical display for clearer reporting. | Allow end‑users to select dimensions at runtime and automatically reshape the pivot layout. | Standardize exported Excel templates by ensuring key fields appear in the row area before distribution.
// AI Prompts: Write C# code with Aspose.Cells that transfers a given pivot field from the Column area to the Row area and updates the pivot. | Explain the sequence of API calls needed to remove a column field, add it as a row field, and recalculate a pivot table in Aspose.Cells. | Provide error‑handling examples for missing source files or empty pivot collections when moving a field between areas.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that loads an Excel workbook, locates the first pivot table, removes a specified column field, adds the same field to the row area, refreshes the pivot cache, recalculates data, and saves the modified file using Aspose.Cells.
class Program
{
    static void Main()
    {
        const string sourcePath = "PivotSource.xlsx";
        const string outputPath = "PivotMoved.xlsx";

        try
        {
            // Verify that the source workbook exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                return;
            }

            // Load the workbook that contains the pivot table
            Workbook workbook = new Workbook(sourcePath);
            Worksheet sheet = workbook.Worksheets[0];

            // Get the first pivot table (adjust index if needed)
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            PivotTable pivotTable = sheet.PivotTables[0];

            // Name of the field that is currently in the Column area and needs to be moved to Row area
            string fieldNameInColumn = "Region"; // <-- change to your actual column field name

            // 1. Remove the field from the Column area
            pivotTable.RemoveField(PivotFieldType.Column, fieldNameInColumn);

            // 2. Add the same field to the Row area (this changes its parent field)
            pivotTable.AddFieldToArea(PivotFieldType.Row, fieldNameInColumn);

            // Refresh and recalculate the pivot table to reflect the changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Pivot table updated and saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
