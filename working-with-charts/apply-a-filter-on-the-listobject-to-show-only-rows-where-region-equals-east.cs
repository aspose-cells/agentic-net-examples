using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class ListObjectRegionFilterDemo
{
    static void Main()
    {
        // Create a new workbook (create rule)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with a "Region" column
        // Header row
        sheet.Cells["A1"].PutValue("Product");
        sheet.Cells["B1"].PutValue("Region");
        sheet.Cells["C1"].PutValue("Sales");

        // Data rows
        sheet.Cells["A2"].PutValue("Laptop");
        sheet.Cells["B2"].PutValue("East");
        sheet.Cells["C2"].PutValue(1200);

        sheet.Cells["A3"].PutValue("Smartphone");
        sheet.Cells["B3"].PutValue("West");
        sheet.Cells["C3"].PutValue(800);

        sheet.Cells["A4"].PutValue("Monitor");
        sheet.Cells["B4"].PutValue("East");
        sheet.Cells["C4"].PutValue(250);

        sheet.Cells["A5"].PutValue("Tablet");
        sheet.Cells["B5"].PutValue("South");
        sheet.Cells["C5"].PutValue(400);

        // Define the range of the table (including header)
        int firstRow = 0;   // zero‑based index for row 1
        int firstCol = 0;   // column A
        int lastRow = 4;    // row 5 (zero‑based)
        int lastCol = 2;    // column C

        // Add a ListObject (table) covering the data range
        int listObjectIndex = sheet.ListObjects.Add(firstRow, firstCol, lastRow, lastCol, true);
        ListObject listObject = sheet.ListObjects[listObjectIndex];

        // Enable AutoFilter for the table
        listObject.HasAutoFilter = true;

        // Apply filter on the "Region" column (index 1) to show only rows where Region = "East"
        // The AutoFilter property is available when HasAutoFilter is true
        listObject.AutoFilter.Filter(1, "East");
        listObject.AutoFilter.Refresh();

        // Save the workbook (save rule)
        workbook.Save("ListObjectRegionFilterDemo.xlsx");
    }
}