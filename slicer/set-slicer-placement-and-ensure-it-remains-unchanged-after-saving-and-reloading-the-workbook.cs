// Title: Set slicer position, size, column layout and style in an Excel workbook using Aspose.Cells for .NET and keep the settings unchanged after saving and reloading
// Common Searches: Aspose.Cells C# how to set slicer top and left pixel coordinates | preserve slicer dimensions after saving workbook with Aspose.Cells .NET | change slicer number of columns and style in a pivot table using Aspose.Cells | verify slicer properties after reloading an Excel file with Aspose.Cells | Aspose.Cells set slicer style type programmatically in C#
// Tags: slicer position configuration Aspose.Cells | slicer dimensions persistence .NET | slicer column count setting Aspose.Cells | apply slicer style Aspose.Cells | pivot table slicer placement verification

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

// The example creates a workbook with sample data, adds a pivot table, inserts a slicer linked to the pivot, configures its TopPixel, LeftPixel, WidthPixel, HeightPixel, NumberOfColumns, and StyleType, saves the file, reloads it, and compares the slicer's properties before and after to confirm that placement, size, layout, and style persist.
class Program
{
    static void Main()
    {
        try
        {
            // -------------------- Create workbook and sample data --------------------
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Header row
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Product");
            cells["C1"].PutValue("Amount");

            // Sample rows
            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue("Apple");
            cells["C2"].PutValue(10);

            cells["A3"].PutValue("Fruit");
            cells["B3"].PutValue("Banana");
            cells["C3"].PutValue(15);

            cells["A4"].PutValue("Vegetable");
            cells["B4"].PutValue("Carrot");
            cells["C4"].PutValue(8);

            // -------------------- Create a pivot table --------------------
            int pivotIndex = worksheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------- Add a slicer linked to the pivot table --------------------
            // Destination cell (where the slicer will be placed) must be a valid cell address.
            int slicerIndex = worksheet.Slicers.Add(pivotTable, "E5", "Category");
            Slicer slicer = worksheet.Slicers[slicerIndex];

            // -------------------- Set slicer placement (position, size, style) --------------------
            slicer.TopPixel = 100;          // distance from the top of the worksheet (in pixels)
            slicer.LeftPixel = 400;         // distance from the left side (in pixels)
            slicer.WidthPixel = 150;        // width of the slicer (in pixels)
            slicer.HeightPixel = 200;       // height of the slicer (in pixels)
            slicer.NumberOfColumns = 1;     // single column layout
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2; // built‑in style

            // -------------------- Save the workbook --------------------
            string filePath = "SlicerPlacementDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------- Reload the workbook and verify placement --------------------
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"The file '{filePath}' was not found after saving.");

            Workbook reloadedWorkbook = new Workbook(filePath);
            Worksheet reloadedWorksheet = reloadedWorkbook.Worksheets[0];

            // Assuming only one slicer exists on the sheet
            if (reloadedWorksheet.Slicers.Count == 0)
                throw new InvalidOperationException("No slicer found in the reloaded worksheet.");

            Slicer reloadedSlicer = reloadedWorksheet.Slicers[0];

            Console.WriteLine("=== Original slicer placement ===");
            PrintSlicerInfo(slicer);

            Console.WriteLine("\n=== Reloaded slicer placement ===");
            PrintSlicerInfo(reloadedSlicer);

            bool placementUnchanged =
                slicer.TopPixel == reloadedSlicer.TopPixel &&
                slicer.LeftPixel == reloadedSlicer.LeftPixel &&
                slicer.WidthPixel == reloadedSlicer.WidthPixel &&
                slicer.HeightPixel == reloadedSlicer.HeightPixel &&
                slicer.NumberOfColumns == reloadedSlicer.NumberOfColumns &&
                slicer.StyleType == reloadedSlicer.StyleType;

            Console.WriteLine($"\nPlacement unchanged after save/load: {placementUnchanged}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    static void PrintSlicerInfo(Slicer s)
    {
        Console.WriteLine($"Name            : {s.Name}");
        Console.WriteLine($"TopPixel        : {s.TopPixel}");
        Console.WriteLine($"LeftPixel       : {s.LeftPixel}");
        Console.WriteLine($"WidthPixel      : {s.WidthPixel}");
        Console.WriteLine($"HeightPixel     : {s.HeightPixel}");
        Console.WriteLine($"NumberOfColumns : {s.NumberOfColumns}");
        Console.WriteLine($"StyleType       : {s.StyleType}");
    }
}
