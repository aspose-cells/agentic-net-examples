using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a pivot table
        sheet.Cells["A1"].PutValue("Fruit");
        sheet.Cells["A2"].PutValue("Apple");
        sheet.Cells["A3"].PutValue("Orange");
        sheet.Cells["A4"].PutValue("Banana");
        sheet.Cells["B1"].PutValue("Sales");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["B4"].PutValue(200);

        // Add a pivot table based on the data
        int pivotIndex = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = sheet.PivotTables[pivotIndex];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
        pivot.RefreshData();
        pivot.CalculateData();

        // Add a slicer linked to the pivot table
        SlicerCollection slicers = sheet.Slicers;
        int slicerIndex = slicers.Add(pivot, "E2", "Fruit");

        // Remove the slicer from the collection
        slicers.RemoveAt(slicerIndex);

        // Save the workbook as an XLSX file
        string filePath = "WorkbookWithoutSlicer.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the saved workbook
        Workbook loadedWorkbook = new Workbook(filePath);

        // Verify that the slicer has been removed
        int remainingSlicers = loadedWorkbook.Worksheets[0].Slicers.Count;
        Console.WriteLine("Remaining slicers after load: " + remainingSlicers);
    }
}