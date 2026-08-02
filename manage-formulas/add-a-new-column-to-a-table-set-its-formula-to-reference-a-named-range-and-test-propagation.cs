using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsTableNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the table (ID, Value, Result)
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Value");
            cells["C1"].PutValue("Result");

            // Add some rows
            cells["A2"].PutValue(1);
            cells["B2"].PutValue(10);
            cells["A3"].PutValue(2);
            cells["B3"].PutValue(20);
            cells["A4"].PutValue(3);
            cells["B4"].PutValue(30);

            // Create a ListObject (table) that includes the three columns
            int tableIndex = sheet.ListObjects.Add("A1", "C4", true);
            ListObject table = sheet.ListObjects[tableIndex];

            // Define a named range that refers to the "Value" column (B2:B4)
            int nameIndex = workbook.Worksheets.Names.Add("MyValues");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // Set the reference using A1 notation; false,false indicate A1 format and global scope
            namedRange.SetRefersTo("=Sheet1!$B$2:$B$4", false, false);

            // Set the formula for the third column ("Result") to reference the named range.
            // Using SUM to produce a scalar value per row; the formula will be applied to each cell in the column.
            ListColumn resultColumn = table.ListColumns[2]; // zero‑based index, third column
            resultColumn.Formula = "=SUM(MyValues)";

            // Calculate formulas so that values are populated
            workbook.CalculateFormula();

            // Verify propagation: output the formula and calculated value for each data row in the "Result" column
            Console.WriteLine("Result column after setting formula referencing named range:");
            for (int row = 1; row <= table.DataRange.RowCount; row++) // data rows start after header
            {
                // Cell address for the current row in the Result column
                Cell resultCell = table.DataRange[row - 1, 2]; // column index 2 within the table's data range
                Console.WriteLine($"{resultCell.Name} -> Formula: {resultCell.Formula}, Value: {resultCell.Value}");
            }

            // Save the workbook
            workbook.Save("TableWithNamedRangeFormula.xlsx");
        }
    }
}