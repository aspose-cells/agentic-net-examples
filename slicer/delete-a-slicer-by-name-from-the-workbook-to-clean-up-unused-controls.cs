using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Tables;

class DeleteSlicerByName
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data and create a table to attach a slicer to
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["A4"].PutValue("A");
        int tableIdx = sheet.ListObjects.Add("A1", "A4", true);
        ListObject table = sheet.ListObjects[tableIdx];

        // Add a slicer and give it a distinct name
        int slicerIdx = sheet.Slicers.Add(table, 0, "C1");
        Slicer slicer = sheet.Slicers[slicerIdx];
        slicer.Name = "TargetSlicer";

        // Retrieve the slicer by its name and remove it from the collection
        SlicerCollection slicers = sheet.Slicers;
        Slicer slicerToRemove = slicers["TargetSlicer"]; // indexer by name
        slicers.Remove(slicerToRemove); // delete the slicer

        // Save the workbook
        workbook.Save("DeletedSlicer.xlsx");
    }
}