using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class DeleteSlicerDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Populate data for a pivot table
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Cells["A1"].PutValue("Fruit");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["A3"].PutValue("Orange");
        dataSheet.Cells["A4"].PutValue("Banana");
        dataSheet.Cells["B1"].PutValue("Quantity");
        dataSheet.Cells["B2"].PutValue(10);
        dataSheet.Cells["B3"].PutValue(5);
        dataSheet.Cells["B4"].PutValue(8);

        // Add a pivot table based on the data
        int pivotIdx = dataSheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
        PivotTable pivot = dataSheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
        pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");

        // Add a new worksheet to host slicers
        Worksheet slicerSheet = workbook.Worksheets.Add("SlicerSheet");

        // Add two slicers linked to the pivot table
        int slicerIdx1 = slicerSheet.Slicers.Add(pivot, "E1", "Fruit");
        Slicer slicer1 = slicerSheet.Slicers[slicerIdx1];
        slicer1.Name = "FruitSlicer1";

        int slicerIdx2 = slicerSheet.Slicers.Add(pivot, "E15", "Fruit");
        Slicer slicer2 = slicerSheet.Slicers[slicerIdx2];
        slicer2.Name = "FruitSlicer2";

        // Delete the second slicer (index 1) using RemoveAt
        slicerSheet.Slicers.RemoveAt(1); // Removes "FruitSlicer2"

        // Save the workbook as an XLSX file
        string filePath = "DeletedSlicerDemo.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the saved workbook to verify the operation
        Workbook loadedWorkbook = new Workbook(filePath);
        Console.WriteLine($"Loaded workbook contains {loadedWorkbook.Worksheets.Count} worksheet(s).");
    }
}