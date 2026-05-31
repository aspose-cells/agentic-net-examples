using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerColorPersistenceDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Prepare sample data for the pivot table
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["A2"].PutValue("Fruit");
            dataSheet.Cells["A3"].PutValue("Fruit");
            dataSheet.Cells["A4"].PutValue("Vegetable");
            dataSheet.Cells["B1"].PutValue("Amount");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["B4"].PutValue(15);

            // Add a pivot table based on the data
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIdx];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);      // Category
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);     // Amount

            // Add a slicer linked to the pivot table
            Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
            int slicerIdx = slicerSheet.Slicers.Add(pivotTable, "A1", "Category");
            Slicer slicer = slicerSheet.Slicers[slicerIdx];

            // Set a built‑in slicer style (this defines the slice colors)
            slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

            // Save the workbook to disk
            string filePath = "SlicerColorPersistenceDemo.xlsx";
            workbook.Save(filePath);

            // Reload the workbook
            Workbook loadedWorkbook = new Workbook(filePath);

            // Retrieve the slicer from the reloaded workbook
            Slicer loadedSlicer = loadedWorkbook.Worksheets["Slicer"].Slicers[0];

            // Verify that the slicer style (and thus its colors) persisted
            bool stylePersisted = loadedSlicer.StyleType == SlicerStyleType.SlicerStyleDark2;
            Console.WriteLine($"Slicer style persisted after reload: {stylePersisted}");
            Console.WriteLine($"Loaded slicer style: {loadedSlicer.StyleType}");
        }
    }
}