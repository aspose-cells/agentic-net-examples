using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

class ValidateSlicerCustomColors
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet dataSheet = workbook.Worksheets[0];

        // Populate sample data for the pivot table
        dataSheet.Cells["A1"].PutValue("Product");
        dataSheet.Cells["A2"].PutValue("Apple");
        dataSheet.Cells["A3"].PutValue("Banana");
        dataSheet.Cells["A4"].PutValue("Orange");
        dataSheet.Cells["B1"].PutValue("Sales");
        dataSheet.Cells["B2"].PutValue(120);
        dataSheet.Cells["B3"].PutValue(80);
        dataSheet.Cells["B4"].PutValue(150);

        // Add a pivot table based on the data
        Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
        int pivotIdx = pivotSheet.PivotTables.Add("A1:B4", "C3", "PivotTable1");
        PivotTable pivot = pivotSheet.PivotTables[pivotIdx];
        pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Product
        pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Sales

        // Add a slicer linked to the pivot table
        Worksheet slicerSheet = workbook.Worksheets.Add("Slicer");
        int slicerIdx = slicerSheet.Slicers.Add(pivot, "A1", "Product");
        Slicer slicer = slicerSheet.Slicers[slicerIdx];

        // Set a built‑in slicer style (acts as a custom visual setting)
        slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

        // OPTIONAL: Change a palette color that could be used by the slicer style
        // (Demonstrates custom color handling in the workbook palette)
        Color customColor = Color.FromArgb(255, 200, 100, 50); // a custom orange tone
        workbook.ChangePalette(customColor, 55); // modify the last palette entry

        // Save the workbook to disk
        string filePath = "SlicerCustomColorDemo.xlsx";
        workbook.Save(filePath);

        // Reload the workbook from disk
        Workbook loadedWb = new Workbook(filePath);
        Worksheet loadedSlicerSheet = loadedWb.Worksheets["Slicer"];
        Slicer loadedSlicer = loadedSlicerSheet.Slicers[0];

        // Validate that the slicer style persisted
        bool stylePersisted = loadedSlicer.StyleType == SlicerStyleType.SlicerStyleDark2;
        Console.WriteLine("Slicer style persisted after reload: " + stylePersisted);

        // Validate that the custom palette color persisted
        bool customColorPersisted = loadedWb.IsColorInPalette(customColor);
        Console.WriteLine("Custom palette color persisted after reload: " + customColorPersisted);
    }
}