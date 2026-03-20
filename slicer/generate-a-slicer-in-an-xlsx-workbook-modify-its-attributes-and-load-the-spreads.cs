using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class SlicerDemo
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        Workbook workbook = new Workbook();                     // create
        Worksheet sheet = workbook.Worksheets[0];

        // Populate worksheet with sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(30);

        // Add a ListObject (table) covering the data range
        int tableIdx = sheet.ListObjects.Add(0, 0, 3, 1, true);
        ListObject table = sheet.ListObjects[tableIdx];

        // Add a slicer linked to the first column of the table
        int slicerIdx = sheet.Slicers.Add(table, 0, "E2");    // add slicer
        Slicer slicer = sheet.Slicers[slicerIdx];

        // Modify slicer attributes
        slicer.Caption = "Category Slicer";
        slicer.NumberOfColumns = 2;
        slicer.LockedPosition = true;
        slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
        slicer.ShowCaption = true;
        slicer.ShowAllItems = true;

        // Save the workbook to disk
        string filePath = "SlicerDemo.xlsx";
        workbook.Save(filePath);                               // save

        // ---------- Load the workbook programmatically ----------
        Workbook loadedWb = new Workbook(filePath);            // load
        Worksheet loadedSheet = loadedWb.Worksheets[0];

        // Access the slicer from the loaded workbook
        Slicer loadedSlicer = loadedSheet.Slicers[0];
        Console.WriteLine("Loaded slicer caption: " + loadedSlicer.Caption);
        Console.WriteLine("Number of columns: " + loadedSlicer.NumberOfColumns);
        Console.WriteLine("Locked position: " + loadedSlicer.LockedPosition);
    }
}