// Title: Create a dynamic array formula that references a ListObject column and capture the #REF! error after deleting the table using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set a dynamic array formula to spill a table column values | C# Aspose.Cells detect #REF! error when a ListObject is removed | How to reference a ListObject column in a formula with Aspose.Cells .NET | Spill array formula from a table and handle error after table deletion in Aspose.Cells | Aspose.Cells calculate formulas after removing a ListObject programmatically
// Tags: dynamic array formula ListObject Aspose.Cells | reference ListObject column in C# formula | handle #REF! after table deletion Aspose.Cells | spill table column values using Aspose.Cells | recalculate workbook after ListObject removal C#

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// The example creates a workbook, adds a single‑column ListObject named MyTable, assigns the dynamic array formula "=MyTable[Value]" to cell C1, forces calculation to spill the column values, prints the spilled data, removes the table, recalculates the workbook, and then checks cell C1 for a #REF! error, outputting the result.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate data for the table (A1 header, A2:A6 values)
            ws.Cells["A1"].PutValue("Value");
            ws.Cells["A2"].PutValue(10);
            ws.Cells["A3"].PutValue(20);
            ws.Cells["A4"].PutValue(30);
            ws.Cells["A5"].PutValue(40);
            ws.Cells["A6"].PutValue(50);

            // Add a ListObject (table) covering A1:A6 and name it "MyTable"
            int firstRow = 0;   // zero‑based index for row 1
            int firstCol = 0;   // zero‑based index for column A
            int totalRows = 6;  // rows 1‑6
            int totalCols = 1;  // single column
            int tableIndex = ws.ListObjects.Add(firstRow, firstCol,
                                                firstRow + totalRows - 1,
                                                firstCol + totalCols - 1, true);
            ListObject table = ws.ListObjects[tableIndex];
            // Use DisplayName to set the table name
            table.DisplayName = "MyTable";

            // Set a dynamic array formula that references the table column.
            // The formula will spill the entire column values into adjacent cells.
            ws.Cells["C1"].Formula = "=MyTable[Value]";

            // Force calculation so the array spills
            wb.CalculateFormula();

            // Display spilled values before deleting the table
            Console.WriteLine("Spilled values before deleting the table:");
            for (int i = 0; i < 5; i++) // 5 data rows
            {
                Console.WriteLine(ws.Cells[i, 2].StringValue); // Column C (index 2)
            }

            // Delete the table
            ws.ListObjects.RemoveAt(0);

            // Recalculate formulas after table deletion
            wb.CalculateFormula();

            // Check the result of the formula cell (C1) for error
            Cell formulaCell = ws.Cells["C1"];
            if (formulaCell.Type == CellValueType.IsError)
            {
                Console.WriteLine($"Formula error after table deletion: {formulaCell.StringValue}");
            }
            else
            {
                Console.WriteLine("Formula did not produce an error after table deletion.");
            }

            // Optional: Save the workbook to verify manually (can be omitted)
            // wb.Save("DynamicArrayFormulaDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
