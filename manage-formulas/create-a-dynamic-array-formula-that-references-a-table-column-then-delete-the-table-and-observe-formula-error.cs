// Title: Aspose.Cells .NET – Dynamic Array Formula Referencing a Table Column, Delete the Table, and Capture the #REF! Error
// Description: Demonstrates how to create a ListObject, apply SetDynamicArrayFormula to a column (e.g., MyTable[Numbers]), calculate and spill the results, then delete the table rows with DeleteOptions.UpdateReference, refresh dynamic array formulas, and retrieve the resulting #REF! error before saving the workbook.
// Keywords: Aspose.Cells | dynamic array formula | SetDynamicArrayFormula | ListObject | table column reference | DeleteRows | DeleteOptions.UpdateReference | RefreshDynamicArrayFormulas | C# | .NET | #REF! error | spreadsheet automation
// Common Searches: Aspose.Cells set dynamic array formula to table column | how to delete a table and keep formula references in Aspose.Cells | RefreshDynamicArrayFormulas after row deletion .NET | dynamic array spill #REF! after table removal Aspose | C# example for ListObject and dynamic array formula
// Developer Intent: Show how a dynamic array formula that points to a table column behaves when the source table is removed, using Aspose.Cells for .NET.
// Use Cases: Create a ListObject, assign a dynamic array formula to its column, and read the spilled values. | Delete the table rows while preserving reference updates via DeleteOptions.UpdateReference. | Refresh dynamic array formulas to expose the #REF! error and persist it in the saved workbook.
// AI Prompts: Generate C# code with Aspose.Cells that sets a dynamic array formula referencing a ListObject column, then deletes the table and captures the #REF! error. | Explain how DeleteOptions.UpdateReference affects dynamic array formulas that depend on a deleted table in Aspose.Cells. | Provide a step‑by‑step guide to refresh dynamic array formulas after removing rows that contain a source table using Aspose.Cells .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a ListObject, apply SetDynamicArrayFormula to a column (e.g., MyTable[Numbers]), calculate and spill the results, then delete the table rows with DeleteOptions.UpdateReference, refresh dynamic array formulas, and retrieve the resulting #REF! error before saving the workbook.
class DynamicArrayTableDeletionDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // Add a header for the table
            cells["A1"].PutValue("Numbers");

            // Populate data for the table (A2:A5)
            for (int i = 0; i < 4; i++)
            {
                cells[i + 1, 0].PutValue(i + 10); // 10,11,12,13
            }

            // Create a table (ListObject) over A1:A5 and give it a display name
            int tableIdx = ws.ListObjects.Add(0, 0, 4, 0, true);
            ListObject table = ws.ListObjects[tableIdx];
            table.DisplayName = "MyTable"; // Set table name (DisplayName works across versions)

            // Set a dynamic array formula that references the table column
            Cell formulaCell = cells["B2"];
            formulaCell.SetDynamicArrayFormula("=MyTable[Numbers]", new FormulaParseOptions(), true);

            // Calculate formulas and refresh dynamic array spills
            wb.CalculateFormula();
            wb.RefreshDynamicArrayFormulas(true);

            // Display the initial spilled values
            Console.WriteLine("Initial dynamic array spill:");
            for (int i = 0; i < 4; i++)
            {
                Console.WriteLine($"B{2 + i}: {cells[1 + i, 1].StringValue}");
            }

            // Delete the rows that contain the table (including the header)
            DeleteOptions delOpts = new DeleteOptions { UpdateReference = true };
            ws.Cells.DeleteRows(0, 5, delOpts); // rows 0‑4

            // Refresh dynamic array formulas after the deletion
            wb.RefreshDynamicArrayFormulas(true);

            // Show the formula and its value after the table has been removed
            Console.WriteLine("\nAfter deleting the table:");
            Console.WriteLine($"Cell B2 formula: {formulaCell.Formula}");
            Console.WriteLine($"Cell B2 value: {formulaCell.StringValue}");

            // Save the workbook
            string outputPath = "DynamicArrayTableDeletionDemo.xlsx";
            wb.Save(outputPath);
            Console.WriteLine($"\nWorkbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
