using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ClearAllTableFilters
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Name");
        sheet.Cells["A2"].PutValue(1);
        sheet.Cells["B2"].PutValue("Alice");
        sheet.Cells["A3"].PutValue(2);
        sheet.Cells["B3"].PutValue("Bob");
        sheet.Cells["A4"].PutValue(3);
        sheet.Cells["B4"].PutValue("Charlie");

        // Add a ListObject (table) which includes an auto‑filter by default
        int tableIndex = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Apply a filter to demonstrate that a filter exists
        table.AutoFilter.AddFilter(1, "Bob");
        table.AutoFilter.Refresh();

        // Remove the auto‑filter, clearing all filters applied to the table
        table.RemoveAutoFilter();

        // Save the workbook
        workbook.Save("ClearAllTableFilters.xlsx");
    }
}