// Title: Aspose.Cells for .NET – Dynamic Array Formula Referencing a Table Column and #REF! After Table Deletion
// Description: This C# example shows how to create a workbook, add a ListObject (table) named MyTable, set a dynamic array formula =MyTable[Values] in cell D1, calculate and spill the results, then delete the table with DeleteOptions.UpdateReference. After refreshing formulas, the spill range displays a #REF! error, demonstrating how Aspose.Cells handles broken structured references.
// Keywords: Aspose.Cells | .NET | C# | dynamic array formula | structured reference | ListObject | Excel table deletion | #REF! error | RefreshDynamicArrayFormulas | DeleteOptions.UpdateReference | spill range | formula recalculation
// Common Searches: Aspose.Cells set dynamic array formula referencing a table column | how to delete a ListObject and update dependent formulas in Aspose.Cells | refresh dynamic array formulas after table removal Aspose.Cells | structured reference #REF! error .NET Aspose.Cells | C# example for dynamic array spill range handling
// Developer Intent: Show how a dynamic array formula that points to a table column behaves when the table is removed, and how to propagate the resulting #REF! error using Aspose.Cells.
// Use Cases: Create a ListObject and assign a dynamic array formula using a structured reference. | Delete a table while automatically updating or invalidating dependent formulas. | Refresh dynamic array formulas to propagate errors and examine the updated spill range.
// AI Prompts: Write C# code with Aspose.Cells that creates a table, sets a dynamic array formula referencing its column, deletes the table, and refreshes formulas to display the #REF! error. | Explain the effect of DeleteOptions.UpdateReference on dynamic array formulas that use structured references in Aspose.Cells. | Provide a step‑by‑step guide to capture spilled values before and after removing a table that a dynamic array formula depends on.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDynamicArrayTableDemo
{
    // This C# example shows how to create a workbook, add a ListObject (table) named MyTable, set a dynamic array formula =MyTable[Values] in cell D1, calculate and spill the results, then delete the table with DeleteOptions.UpdateReference. After refreshing formulas, the spill range displays a #REF! error, demonstrating how Aspose.Cells handles broken structured references.
    class Program
    {
        static void Main()
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet sheet = wb.Worksheets[0];
                Cells cells = sheet.Cells;

                // 2. Populate sample data in column B (B2:B4)
                cells["B2"].PutValue(10);
                cells["B3"].PutValue(20);
                cells["B4"].PutValue(30);

                // 3. Create a table (ListObject) that includes the header in B1 and the data in B2:B4
                cells["B1"].PutValue("Values"); // header
                int tableIndex = sheet.ListObjects.Add(0, 1, 3, 1, true); // rows 0‑3, column 1 (B)
                ListObject table = sheet.ListObjects[tableIndex];
                // Set a stable name for the structured reference
                table.DisplayName = "MyTable";

                // 4. Set a dynamic array formula that references the table column
                // The structured reference syntax is TableName[ColumnName]
                Cell formulaCell = cells["D1"];
                string dynArrayFormula = $"=MyTable[Values]";
                formulaCell.SetDynamicArrayFormula(dynArrayFormula, new FormulaParseOptions(), true);

                // 5. Calculate formulas and refresh dynamic array spills
                wb.CalculateFormula();
                wb.RefreshDynamicArrayFormulas(true);

                // 6. Output the spilled values (D1:D3) before deleting the table
                Console.WriteLine("Spill range before deleting the table:");
                for (int row = 0; row < 3; row++)
                {
                    Console.WriteLine($"D{row + 1}: {cells[row, 3].Value}");
                }

                // 7. Delete the table rows (including header) with reference update
                DeleteOptions delOpts = new DeleteOptions { UpdateReference = true };
                int firstRow = table.StartRow; // header row index
                int rowsToDelete = table.EndRow - table.StartRow + 1; // total rows of the table
                cells.DeleteRows(firstRow, rowsToDelete, delOpts);

                // 8. Refresh dynamic array formulas again so the #REF! error propagates
                wb.RefreshDynamicArrayFormulas(true);

                // 9. Show the formula and value after the table has been removed
                Console.WriteLine("\nAfter deleting the table:");
                Console.WriteLine($"Formula in D1: {formulaCell.Formula}");
                Console.WriteLine($"Value in D1 (should be error): {formulaCell.Value}");

                // 10. Save the workbook for inspection
                wb.Save("DynamicArrayTableDeletionDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
