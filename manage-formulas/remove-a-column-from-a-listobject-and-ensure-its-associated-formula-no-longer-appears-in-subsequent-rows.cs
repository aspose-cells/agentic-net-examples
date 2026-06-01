using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class RemoveListObjectColumn
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("Item");
        sheet.Cells["B1"].PutValue("Quantity");
        sheet.Cells["C1"].PutValue("Total");

        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["B2"].PutValue(2);
        sheet.Cells["A3"].PutValue("Banana");
        sheet.Cells["B3"].PutValue(3);

        // Set a formula for the Total column (C) that depends on Quantity (B)
        sheet.Cells["C2"].Formula = "=B2*5";
        sheet.Cells["C3"].Formula = "=B3*5";

        // Create a ListObject (table) that includes the data range A1:C3
        int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Assign the same formula to the ListColumn via the ListColumn property
        // This ensures the formula is applied to all rows of the column within the table
        ListColumn totalColumn = table.ListColumns[2]; // Column C (zero‑based index)
        totalColumn.Formula = "=B2*5";

        // Before deleting, clear the ListColumn formula so it does not persist after the column is removed
        totalColumn.Formula = string.Empty;

        // Delete the column (index 2 corresponds to column C) and update references in formulas
        sheet.Cells.DeleteColumn(2, true);

        // Save the modified workbook
        workbook.Save("RemovedColumn.xlsx");
    }
}