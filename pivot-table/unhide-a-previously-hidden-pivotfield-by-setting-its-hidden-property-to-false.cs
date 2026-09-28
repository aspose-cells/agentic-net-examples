// Title: How to unhide all items of a hidden row field in an Excel pivot table using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, accesses the first pivot table's first row field, and makes every hidden item visible before saving the file. | Provide a C# example that loops through a PivotField's items, calls HideItem(index, false) for each, refreshes the pivot table, and writes the updated workbook to disk. | Show how to programmatically change a PivotField's visibility to false for all items using Aspose.Cells and ensure the pivot table recalculates.
// Common Searches: Aspose.Cells C# unhide pivot table row field items programmatically | How to make hidden items visible in an Excel pivot table with Aspose.Cells .NET | Refresh pivot table after changing field visibility using Aspose.Cells C#
// Tags: Aspose.Cells hideitem false | C# pivot table row field visibility | refresh pivot table after modification | load workbook modify pivot field Aspose.Cells | iterate pivotfield items C#

using Aspose.Cells;
using Aspose.Cells.Pivot;
using System;
using System.IO;

// The sample loads PivotTableHiddenField.xlsx, retrieves the first pivot table's first row field, iterates over all its items calling HideItem(i, false) to unhide them, refreshes and recalculates the pivot table, and saves the result as PivotTableUnhiddenField.xlsx.
class UnhidePivotField
{
    static void Main()
    {
        const string inputPath = "PivotTableHiddenField.xlsx";
        const string outputPath = "PivotTableUnhiddenField.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook containing the hidden pivot items
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one pivot table
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            // Access the first pivot table
            PivotTable pivotTable = sheet.PivotTables[0];

            // Ensure the pivot table has at least one row field
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("Pivot table has no row fields.");
                return;
            }

            // Choose the first row field to unhide its items
            PivotField pivotField = pivotTable.RowFields[0];

            // Unhide all items in the selected field
            for (int i = 0; i < pivotField.ItemCount; i++)
            {
                pivotField.HideItem(i, false);
            }

            // Refresh and recalculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
