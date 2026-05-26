using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class InsertRowInSalesTable
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Create sample data for the table -----
        // Header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["C1"].PutValue("Double");

        // Existing data rows
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue(20);

        // ----- Create a ListObject (table) that covers the data -----
        // The range "A1:C3" includes the header and the two data rows
        int tableIdx = sheet.ListObjects.Add("A1", "C3", true);
        ListObject salesTable = sheet.ListObjects[tableIdx];

        // ----- Set the formula for the existing rows -----
        // Row offset 1 = second row in the table (sheet row 2)
        salesTable.PutCellFormula(1, 2, "=B2*2"); // C2 = B2 * 2
        // Row offset 2 = third row in the table (sheet row 3)
        salesTable.PutCellFormula(2, 2, "=B3*2"); // C3 = B3 * 2

        // ----- Insert a new row into the table -----
        // Row offset 3 = fourth row in the table (sheet row 4)
        salesTable.PutCellValue(3, 0, 3);   // ID = 3
        salesTable.PutCellValue(3, 1, 30);  // Value = 30
        // Formula for the new row references its own Value cell (B4)
        salesTable.PutCellFormula(3, 2, "=B4*2"); // C4 = B4 * 2

        // Recalculate all formulas so the new row's result is computed
        workbook.CalculateFormula();

        // ----- Verify that the formula was applied correctly -----
        Console.WriteLine("New row calculated Double value: " + sheet.Cells["C4"].Value); // Expected: 60

        // Save the workbook (creation and saving follow the required lifecycle)
        workbook.Save("SalesTableDemo.xlsx");
    }
}