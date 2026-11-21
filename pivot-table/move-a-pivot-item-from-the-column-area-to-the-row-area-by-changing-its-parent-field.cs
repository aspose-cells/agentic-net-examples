// Title: Move a pivot table column field to the row area programmatically with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells move pivot column field to row field | how to change pivot table field orientation using Aspose.Cells .NET | reassign a pivot item from column area to row area with Aspose.Cells | Aspose.Cells example for removing a column field and adding it to rows | programmatically transfer a pivot field from columns to rows in Excel using C#
// Tags: Aspose.Cells pivot field relocation | C# modify pivot table layout | Excel pivot column-to-row conversion Aspose.Cells | programmatic pivot area adjustment .NET | Aspose.Cells change pivot field orientation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;   // Required for PivotTable and PivotField classes

// The sample loads an existing Excel workbook, checks for a pivot table on the first worksheet, removes the first column field from the pivot's ColumnFields collection, adds the same field to the RowFields collection, and saves the updated workbook.
class PivotItemMover
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("Error: No pivot tables found on the first worksheet.");
                return;
            }

            // Get the first pivot table on the worksheet (adjust index if needed)
            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure there is at least one column field to move
            if (pivotTable.ColumnFields.Count == 0)
            {
                Console.WriteLine("Error: The pivot table does not contain any column fields to move.");
                return;
            }

            // Retrieve the first column field
            PivotField columnField = pivotTable.ColumnFields[0];

            // Remove the field from the column area
            pivotTable.ColumnFields.RemoveAt(0);

            // Add the same field to the row area
            pivotTable.RowFields.Add(columnField);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
