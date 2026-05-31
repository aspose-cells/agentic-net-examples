using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class SetTableShowTotalsFalse
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the table
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["B3"].PutValue(15);
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B4"].PutValue(8);

        // Create a ListObject (table) that includes the data range
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // (Optional) Show the totals row first to demonstrate the change
        table.ShowTotals = true;

        // Hide the totals row by setting ShowTotals to false
        table.ShowTotals = false;

        // Save the workbook to a file
        workbook.Save("TableWithoutTotals.xlsx");
    }
}