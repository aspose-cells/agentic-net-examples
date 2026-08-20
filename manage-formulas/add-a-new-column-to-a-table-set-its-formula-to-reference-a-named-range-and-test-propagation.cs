// Title: Add a Column with a Named‑Range Formula to an Aspose.Cells Table and Auto‑Propagate on New Rows (C#)
// Description: Demonstrates how to create a workbook, define a named range, build a ListObject, insert a new column, resize the table, assign a =SUM(MyRange) formula to the column, add a row, recalculate formulas, and output the calculated values to confirm automatic propagation.
// Keywords: Aspose.Cells | C# | ListObject | add column to table | named range formula | formula propagation | resize table | calculate formulas | Excel automation | dynamic column
// Common Searches: Aspose.Cells add column to ListObject with formula | set named range formula in Aspose.Cells table column | auto‑propagate table formulas after inserting rows Aspose.Cells | resize Aspose.Cells table after inserting a column | C# example named range formula in Excel table
// Developer Intent: Add a new column to an existing Aspose.Cells table, apply a formula that references a named range, and ensure the formula automatically fills cells of rows added later.
// Use Cases: Create a calculated column that always sums a predefined range, regardless of data growth. | Add a dynamic summary column that updates instantly when new records are appended. | Centralize range references with named ranges for easier maintenance across multiple table columns.
// AI Prompts: Generate C# code using Aspose.Cells to insert a column into a ListObject, set its formula to =SUM(MyRange), and recalculate the workbook. | Explain how to resize an Aspose.Cells table after inserting a column and ensure the new column’s formula propagates to newly added rows. | Provide troubleshooting steps when a named‑range formula does not auto‑fill after adding rows to an Aspose.Cells table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook, define a named range, build a ListObject, insert a new column, resize the table, assign a =SUM(MyRange) formula to the column, add a row, recalculate formulas, and output the calculated values to confirm automatic propagation.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the initial table (two columns: ID and Value)
            ws.Cells["A1"].PutValue("ID");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue(1);
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue(2);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue(3);
            ws.Cells["B4"].PutValue(30);

            // Create a named range "MyRange" that refers to the Value column (B2:B4)
            int nameIdx = wb.Worksheets.Names.Add("MyRange");
            Name myRange = wb.Worksheets.Names[nameIdx];
            myRange.RefersTo = "=Sheet1!$B$2:$B$4";

            // Add a ListObject (table) covering the existing data range A1:B4
            int tblIdx = ws.ListObjects.Add("A1", "B4", true);
            ListObject table = ws.ListObjects[tblIdx];
            // ShowHeaders is true by default; no explicit property needed

            // Insert a new column after the existing ones (column index 2 corresponds to column C)
            ws.Cells.InsertColumn(2, true);

            // Expand the table to include the newly inserted column (range A1:C4)
            // Resize(startRow, startColumn, totalRows, totalColumns, expandRows)
            table.Resize(0, 0, 4, 3, true);

            // Set header for the new column
            ws.Cells["C1"].PutValue("Calc");

            // Set the formula for the new column to reference the named range.
            // This formula will be applied to every cell in the column.
            ListColumn calcColumn = table.ListColumns[2]; // zero‑based index
            calcColumn.Formula = "=SUM(MyRange)";

            // Add a new row to the table; the formula should automatically propagate to the new cell.
            table.PutCellValue(4, 0, 4);   // ID for the new row
            table.PutCellValue(4, 1, 40);  // Value for the new row

            // Recalculate all formulas in the workbook
            wb.CalculateFormula();

            // Output the values of the Calc column to verify propagation
            Console.WriteLine("Calc column values after adding a row:");
            for (int row = 1; row <= 5; row++) // rows 1..5 (including header row)
            {
                Console.WriteLine($"Row {row}: {ws.Cells[row, 2].Value}");
            }

            // Save the workbook
            wb.Save("TableWithNamedRangeFormula.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
