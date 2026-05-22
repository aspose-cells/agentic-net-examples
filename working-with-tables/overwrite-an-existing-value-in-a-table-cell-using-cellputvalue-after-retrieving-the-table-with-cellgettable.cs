using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class OverwriteTableCell
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Populate sample data that will become a table (range A1:C3)
        cells["A1"].PutValue("ID");
        cells["B1"].PutValue("Name");
        cells["C1"].PutValue("Score");
        cells["A2"].PutValue(1);
        cells["B2"].PutValue("Alice");
        cells["C2"].PutValue(85);
        cells["A3"].PutValue(2);
        cells["B3"].PutValue("Bob");
        cells["C3"].PutValue(90);

        // Create a ListObject (table) covering the populated range
        int tableIndex = sheet.ListObjects.Add(0, 0, 2, 2, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Select a cell inside the table that we want to overwrite (C2)
        Cell targetCell = cells["C2"]; // Current value is 85

        // Retrieve the table that contains this cell using Cell.GetTable()
        ListObject parentTable = targetCell.GetTable();

        // Ensure the cell belongs to the expected table before overwriting
        if (parentTable != null && parentTable == table)
        {
            // Overwrite the existing value using Cell.PutValue
            targetCell.PutValue(95); // New score value
        }

        // Save the workbook
        workbook.Save("OverwriteTableCell.xlsx", SaveFormat.Xlsx);
    }
}