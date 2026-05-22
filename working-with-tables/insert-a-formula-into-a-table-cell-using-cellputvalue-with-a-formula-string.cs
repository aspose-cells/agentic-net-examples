using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class InsertFormulaIntoTableCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        Cells cells = ws.Cells;

        // Add header row
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Value");
        cells["C1"].PutValue("Formula");

        // Add some sample data
        cells["A2"].PutValue(1);
        cells["B2"].PutValue(10);
        cells["A3"].PutValue(2);
        cells["B3"].PutValue(20);

        // Create a table that spans A1:C3
        int tableIndex = ws.ListObjects.Add("A1", "C3", true);
        ListObject table = ws.ListObjects[tableIndex];

        // Append a new row to the table (row offset 3, because header is row 0)
        table.PutCellValue(3, 0, 3);   // ID column
        table.PutCellValue(3, 1, 30);  // Value column

        // Get the cell object for the Formula column in the newly added row
        // Row index 3 (fourth row), column index 2 (C column)
        Cell formulaCell = cells[3, 2];

        // Insert a formula using PutValue with a formula string.
        // The string starts with '=' so Excel treats it as a formula.
        formulaCell.PutValue("=B4*2"); // B4 corresponds to the Value column of the new row

        // Recalculate the workbook so the formula result is evaluated
        wb.CalculateFormula();

        // Display the calculated result
        Console.WriteLine("Result in C4 (Formula column): " + formulaCell.Value);

        // Save the workbook
        wb.Save("TableWithFormula.xlsx");
    }
}