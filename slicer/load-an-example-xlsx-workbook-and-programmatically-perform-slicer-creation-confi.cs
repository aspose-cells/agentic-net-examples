using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class SlicerDemo
{
    static void Main()
    {
        // Load an existing workbook if it exists; otherwise create a new one.
        string inputPath = "example.xlsx";
        Workbook workbook = System.IO.File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

        // Work with the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Ensure there is sample data for a table.
        if (sheet.Cells.MaxDataRow == 0 && sheet.Cells.MaxDataColumn == 0)
        {
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
        }

        // Add a ListObject (table) that covers the populated range.
        int tableIndex = sheet.ListObjects.Add(0, 0, sheet.Cells.MaxDataRow, sheet.Cells.MaxDataColumn, true);
        ListObject table = sheet.ListObjects[tableIndex];

        // Create a slicer for the first column of the table, placing it at cell E2.
        int slicerIndex = sheet.Slicers.Add(table, 0, "E2");
        Slicer slicer = sheet.Slicers[slicerIndex];

        // Configure slicer appearance and behavior.
        slicer.Caption = "Category Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.NumberOfColumns = 2;
        slicer.LockedPosition = true;
        slicer.ShowAllItems = true;

        // Add new data to the worksheet to demonstrate slicer refresh.
        sheet.Cells["A5"].PutValue("C");
        sheet.Cells["B5"].PutValue(40);

        // Refresh the slicer so it reflects the updated data source.
        slicer.Refresh();

        // Save the modified workbook.
        workbook.Save("SlicersDemo.xlsx");
    }
}