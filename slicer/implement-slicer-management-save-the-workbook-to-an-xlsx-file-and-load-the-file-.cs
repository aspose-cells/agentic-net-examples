using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class SlicerDemo
{
    static void Main()
    {
        // Create a new workbook (uses Workbook() constructor)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet ws = workbook.Worksheets[0];

        // Populate sample data for a table
        ws.Cells["A1"].PutValue("Category");
        ws.Cells["B1"].PutValue("Value");
        ws.Cells["A2"].PutValue("A");
        ws.Cells["B2"].PutValue(10);
        ws.Cells["A3"].PutValue("B");
        ws.Cells["B3"].PutValue(20);
        ws.Cells["A4"].PutValue("A");
        ws.Cells["B4"].PutValue(30);
        ws.Cells["A5"].PutValue("B");
        ws.Cells["B5"].PutValue(40);

        // Add a ListObject (table) covering the data range A1:B5
        int tableIdx = ws.ListObjects.Add(0, 0, 4, 1, true);
        ListObject table = ws.ListObjects[tableIdx];

        // Add a slicer for the first column of the table at cell C1
        int slicerIdx = ws.Slicers.Add(table, table.ListColumns[0], "C1");
        Slicer slicer = ws.Slicers[slicerIdx];
        slicer.Caption = "Category Slicer";

        // Save the workbook to an XLSX file (uses Workbook.Save(string, SaveFormat))
        string filePath = "SlicerDemo.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the workbook back from the saved file (uses Workbook(string) constructor)
        Workbook loadedWorkbook = new Workbook(filePath);

        // Access the slicer in the loaded workbook and refresh it
        Worksheet loadedWs = loadedWorkbook.Worksheets[0];
        if (loadedWs.Slicers.Count > 0)
        {
            Slicer loadedSlicer = loadedWs.Slicers[0];
            loadedSlicer.Refresh(); // Refreshes slicer and its underlying PivotTable if any
        }

        // Save the loaded workbook again to verify that the slicer is still functional
        loadedWorkbook.Save("SlicerDemo_Loaded.xlsx", SaveFormat.Xlsx);
    }
}