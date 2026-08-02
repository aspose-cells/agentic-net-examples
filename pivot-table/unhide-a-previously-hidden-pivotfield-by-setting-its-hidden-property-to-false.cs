// Title: Aspose.Cells for .NET – Unhide PivotField Items and Refresh the Pivot Table (C#)
// Description: Loads an existing workbook, accesses the first worksheet and its first pivot table, selects a row PivotField, makes every hidden item visible with HideItem(..., false), refreshes and recalculates the pivot table, then saves the updated file.
// Keywords: Aspose.Cells | C# | PivotTable | PivotField | HideItem | unhide items | show hidden pivot field | refresh pivot table | calculate pivot data | Excel automation | Aspose.Cells for .NET
// Common Searches: how to unhide pivot field items using Aspose.Cells C# | Aspose.Cells hideitem false example | programmatically show hidden rows in a pivot table | refresh pivot table after changing hidden state Aspose | C# code to make all pivot field items visible
// Developer Intent: Reveal hidden items in a PivotField and update the workbook programmatically.
// Use Cases: Generate a complete report by ensuring all pivot field items are visible before exporting. | Batch‑process multiple workbooks to reset hidden states of pivot fields for end‑user consistency. | Prepare an Excel file for PDF conversion where hidden pivot items must be displayed.
// AI Prompts: Write C# code with Aspose.Cells that unhides every item in a specified PivotField and refreshes the pivot table. | Create a reusable method that accepts a worksheet name and a PivotField index, then makes all items visible using Aspose.Cells. | Show how to iterate through all pivot tables in a workbook and unhide hidden row fields with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace AsposeCellsPivotUnhideField
{
    // Loads an existing workbook, accesses the first worksheet and its first pivot table, selects a row PivotField, makes every hidden item visible with HideItem(..., false), refreshes and recalculates the pivot table, then saves the updated file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains a pivot table
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Assume the first pivot table is the target
            if (sheet.PivotTables.Count == 0)
            {
                Console.WriteLine("No pivot tables found in the worksheet.");
                return;
            }

            PivotTable pivotTable = sheet.PivotTables[0];

            // Choose the pivot field to unhide.
            // Here we take the first row field; modify the index or type as required.
            if (pivotTable.RowFields.Count == 0)
            {
                Console.WriteLine("No row fields found in the pivot table.");
                return;
            }

            PivotField pivotField = pivotTable.RowFields[0];

            // Unhide all items in the selected pivot field.
            // The HideItem method sets the hidden state for a specific item.
            // Passing 'false' makes the item visible.
            for (int i = 0; i < pivotField.ItemCount; i++)
            {
                pivotField.HideItem(i, false);
            }

            // Refresh and recalculate the pivot table to apply changes
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the modified workbook
            workbook.Save("output.xlsx");

            Console.WriteLine("Pivot field items have been unhidden and workbook saved as 'output.xlsx'.");
        }
    }
}
