using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with formulas
        sheet.Cells["A1"].PutValue("Qty");
        sheet.Cells["B1"].PutValue("Price");
        sheet.Cells["C1"].PutValue("Total");

        sheet.Cells["A2"].PutValue(2);
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].Formula = "=A2*B2";

        sheet.Cells["A3"].PutValue(5);
        sheet.Cells["B3"].PutValue(7);
        sheet.Cells["C3"].Formula = "=A3*B3";

        // Add a ListObject (table) that covers the range A1:C3
        int tableIndex = sheet.ListObjects.Add("A1", "C3", true);
        ListObject listObject = sheet.ListObjects[tableIndex];

        // Convert the ListObject back to a regular range.
        // This preserves all existing formulas in the cells.
        listObject.ConvertToRange();

        // Save the workbook
        workbook.Save("ListObjectToRange.xlsx");
    }
}