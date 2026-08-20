// Title: Add a Row to an Aspose.Cells ListObject and Auto‑Recalculate Column Formula (C#)
// Description: Demonstrates how to create a workbook with a ListObject (table), assign a formula to the Total column, insert a new data row, trigger CalculateFormula, and verify that the formula updates automatically before saving the file.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Insert row | Table formula | CalculateFormula | auto recalculate | sales table example | programmatic Excel
// Common Searches: Aspose.Cells add row to ListObject C# | how to keep table formulas after inserting rows Aspose.Cells | auto recalculate column formula in Aspose.Cells table | C# insert row into Excel table using Aspose.Cells | Aspose.Cells CalculateFormula after adding rows
// Developer Intent: Insert a new data row into a ListObject and have the existing column formula compute automatically.
// Use Cases: Appending daily sales records while the Total column updates without manual edits. | Expanding a financial ledger programmatically and ensuring derived fields (tax, discount) recalculate. | Generating dynamic reports where rows are added in code and all summary formulas refresh instantly.
// AI Prompts: Write C# code that adds a row to an Aspose.Cells ListObject and automatically applies the existing Total column formula. | Show how to call workbook.CalculateFormula after inserting rows so that new and existing formulas are evaluated. | Provide an example that verifies the calculated value of a formula in a newly inserted row of an Aspose.Cells table.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

// Demonstrates how to create a workbook with a ListObject (table), assign a formula to the Total column, insert a new data row, trigger CalculateFormula, and verify that the formula updates automatically before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Define table headers
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Amount");
        cells["C1"].PutValue("Total");

        // Add initial data rows
        cells["A2"].PutValue(1);
        cells["B2"].PutValue(100);
        cells["A3"].PutValue(2);
        cells["B3"].PutValue(150);

        // Create a ListObject (table) that spans A1:C3
        int tableIdx = sheet.ListObjects.Add("A1", "C3", true);
        ListObject salesTable = sheet.ListObjects[tableIdx];

        // Set the formula for the Total column of existing rows
        // Row offset 1 = first data row (row 2 in the sheet)
        salesTable.PutCellFormula(1, 2, "=B2*2");
        // Row offset 2 = second data row (row 3 in the sheet)
        salesTable.PutCellFormula(2, 2, "=B3*2");

        // Insert a new row into the table using PutCellValue / PutCellFormula
        // Row offset 3 = third data row (row 4 in the sheet)
        salesTable.PutCellValue(3, 0, 3);          // ID
        salesTable.PutCellValue(3, 1, 250);        // Amount
        salesTable.PutCellFormula(3, 2, "=B4*2"); // Total (formula references the new Amount cell)

        // Recalculate all formulas so the Total column updates automatically
        workbook.CalculateFormula();

        // Confirm the formula result for the newly added row
        Console.WriteLine("New row Total (C4): " + sheet.Cells["C4"].Value); // Expected: 500

        // Save the workbook
        workbook.Save("SalesTable.xlsx");
    }
}
