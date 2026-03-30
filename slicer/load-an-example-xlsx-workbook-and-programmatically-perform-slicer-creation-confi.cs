using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;

class SlicerDemo
{
    static void Main()
    {
        string inputPath = "example.xlsx";
        Workbook workbook = System.IO.File.Exists(inputPath) ? new Workbook(inputPath) : new Workbook();

        Worksheet sheet = workbook.Worksheets[0];

        // ✅ Ensure valid sample data (prevents invalid range issues)
        if (sheet.Cells.MaxDataRow < 1)
        {
            sheet.Cells.Clear();

            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);

            sheet.Cells["A4"].PutValue("A");
            sheet.Cells["B4"].PutValue(30);
        }

        // ✅ Safe fixed range for table
        int startRow = 0;
        int startCol = 0;
        int endRow = 3;   // rows 0–3 (A1:A4)
        int endCol = 1;   // columns A–B

        ListObject table;

        // ✅ Avoid duplicate table creation
        if (sheet.ListObjects.Count == 0)
        {
            int tableIndex = sheet.ListObjects.Add(startRow, startCol, endRow, endCol, true);
            table = sheet.ListObjects[tableIndex];
        }
        else
        {
            table = sheet.ListObjects[0];
        }

        // ✅ Avoid duplicate slicer creation
        Slicer slicer;
        if (sheet.Slicers.Count == 0)
        {
            int slicerIndex = sheet.Slicers.Add(table, 0, "E2");
            slicer = sheet.Slicers[slicerIndex];
        }
        else
        {
            slicer = sheet.Slicers[0];
        }

        // ✅ Configure slicer
        slicer.Caption = "Category Slicer";
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.NumberOfColumns = 2;
        slicer.LockedPosition = true;
        slicer.ShowAllItems = true;

        // ✅ Add new data
        int newRow = endRow + 1;
        sheet.Cells[newRow, 0].PutValue("C");
        sheet.Cells[newRow, 1].PutValue(40);

        // ✅ Resize table correctly (FIXED ERROR HERE)
        table.Resize(startRow, startCol, newRow, endCol, true);

        // ✅ Refresh slicer
        slicer.Refresh();

        // ✅ Save file
        workbook.Save("SlicersDemo.xlsx");

        Console.WriteLine("Slicer demo executed successfully.");
    }
}