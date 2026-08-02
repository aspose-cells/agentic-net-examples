// Title: Insert a Formula into an Aspose.Cells Table Cell using Cell.PutValue (C# .NET)
// Description: Demonstrates how to create a workbook, add a ListObject (table) covering A1:C4, locate the underlying worksheet cell for a specific data row, and assign a formula string (e.g., "=B5*2") with Cell.PutValue. The sample prints the formula and saves the file as TableWithFormula.xlsx.
// Keywords: Aspose.Cells C# | Cell.PutValue formula | Aspose.Cells ListObject | table formula .NET | insert formula into table cell | Aspose.Cells sample code | C# Excel table API | Aspose.Cells GitHub example | dynamic formula in ListObject | Aspose.Cells workbook manipulation
// Common Searches: how to set a formula in an Aspose.Cells table column | Cell.PutValue with formula string Aspose.Cells | C# add formula to ListObject cell | Aspose.Cells insert formula into table row | Aspose.Cells table totals row formula example
// Developer Intent: Add a formula to a specific cell inside an Aspose.Cells ListObject using Cell.PutValue.
// Use Cases: Create a calculated column that doubles another column’s value for each row in a table. | Programmatically update formulas when new rows are appended to a ListObject. | Apply row‑specific formulas without rebuilding the worksheet or table structure.
// AI Prompts: Provide C# code that inserts a formula into the "Formula" column of the last data row of an Aspose.Cells ListObject using Cell.PutValue. | Show an example that creates a table, calculates the correct row/column offsets, and sets a formula string with PutValue. | Explain how to target a cell inside an Aspose.Cells table and assign a dynamic formula programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsFormulaInTable
{
    // Demonstrates how to create a workbook, add a ListObject (table) covering A1:C4, locate the underlying worksheet cell for a specific data row, and assign a formula string (e.g., "=B5*2") with Cell.PutValue. The sample prints the formula and saves the file as TableWithFormula.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Add header row for the table
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("Formula");

            // Add some sample data rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue(10);
            cells["A3"].PutValue(2);
            cells["B3"].PutValue(20);
            cells["A4"].PutValue(3);
            cells["B4"].PutValue(30);

            // Define the range that will become a table (A1:C4)
            int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Insert a formula into the "Formula" column of the last data row
            // Using Cell.PutValue with a formula string (starts with '=').
            // The cell is C5 (row index 4, column index 2) because the table adds a totals row by default.
            // To target the data row, we use row offset 3 (zero‑based) within the table.
            // Row offset 3 corresponds to the fourth data row (A4:B4) in the worksheet.
            // Column offset 2 points to the "Formula" column.
            // Retrieve the underlying worksheet cell and set the formula via PutValue.
            Cell targetCell = sheet.Cells[table.StartRow + 3, table.StartColumn + 2];
            targetCell.PutValue("=B5*2"); // Formula: double the value in column B of the same row

            // Optionally, verify that the cell now contains a formula
            Console.WriteLine($"Cell {targetCell.Name} formula: {targetCell.Formula}");

            // Save the workbook
            workbook.Save("TableWithFormula.xlsx");
        }
    }
}
