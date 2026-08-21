// Title: Add a Formula to a Table Cell with Cell.PutValue in Aspose.Cells for .NET
// Description: This example creates a workbook, defines a header and data rows, builds a ListObject (table) over range A1:C3, appends a new row, and uses Cell.PutValue with a formula string ("=B3*2") to insert a calculated value into the "Formula" column before saving the file.
// Keywords: Aspose.Cells C# formula | Cell.PutValue formula string | ListObject add formula | Excel table formula Aspose | programmatic Excel formula .NET | Aspose.Cells table column calculation
// Common Searches: how to set a formula in an Aspose.Cells table cell | Cell.PutValue formula in ListObject C# | add calculated column to Aspose.Cells table | Aspose.Cells insert formula programmatically | C# Aspose.Cells table row formula
// Developer Intent: Insert a calculated formula into a specific cell of a ListObject using Cell.PutValue.
// Use Cases: Create a dynamic "Formula" column after adding new rows to a table. | Reference other columns in the same row with a row‑level formula. | Generate reports where each table row contains a custom calculation defined in code.
// AI Prompts: Show C# code that uses Cell.PutValue to assign a formula to a newly added row in an Aspose.Cells ListObject. | Give an example of applying the same formula to every row of a table with Aspose.Cells for .NET. | Explain how Aspose.Cells interprets a string passed to PutValue as a formula.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsFormulaInTable
{
    // This example creates a workbook, defines a header and data rows, builds a ListObject (table) over range A1:C3, appends a new row, and uses Cell.PutValue with a formula string ("=B3*2") to insert a calculated value into the "Formula" column before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add header row
            sheet.Cells["A1"].PutValue("ID");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["C1"].PutValue("Formula");

            // Add some sample data rows
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);

            // Create a table (ListObject) that spans the data range
            int tableIndex = sheet.ListObjects.Add("A1", "C3", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Add a new row to the table and set values for ID and Value columns
            table.PutCellValue(4, 0, 3);   // Row offset 4, Column offset 0 -> ID
            table.PutCellValue(4, 1, 30);  // Row offset 4, Column offset 1 -> Value

            // Insert a formula into the "Formula" column of the same row using Cell.PutValue
            // The cell coordinates correspond to row index 4 (zero‑based) and column index 2
            Cell formulaCell = sheet.Cells[4, 2];
            formulaCell.PutValue("=B3*2"); // Formula string; Aspose.Cells will treat it as a formula

            // Save the workbook
            workbook.Save("TableWithFormula.xlsx");
        }
    }
}
