using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace SlicerPlacementValidation
{
    class Program
    {
        static void Main()
        {
            // -------------------- Create workbook and data --------------------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data for a pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Sales";
            cells["A2"].Value = "Apple";
            cells["B2"].Value = 120;
            cells["A3"].Value = "Orange";
            cells["B3"].Value = 150;
            cells["A4"].Value = "Banana";
            cells["B4"].Value = 90;

            // -------------------- Create pivot table --------------------
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------- Add slicer linked to the pivot table --------------------
            // Destination cell for slicer upper‑left corner is E1
            int slicerIdx = sheet.Slicers.Add(pivot, "E1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Set explicit placement coordinates via the underlying Shape object
            SlicerShape shape = slicer.Shape;
            shape.Left = 100;   // pixels from left column
            shape.Top = 50;     // pixels from top row
            shape.Width = 200;  // pixels
            shape.Height = 150; // pixels

            // Store original coordinates for later comparison
            int originalLeft = shape.Left;
            int originalTop = shape.Top;
            int originalWidth = shape.Width;
            int originalHeight = shape.Height;

            // -------------------- Save workbook --------------------
            string filePath = "SlicerPlacementDemo.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------- Load workbook --------------------
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Slicer loadedSlicer = loadedSheet.Slicers[slicerIdx];
            SlicerShape loadedShape = loadedSlicer.Shape;

            // Retrieve coordinates after reload
            int loadedLeft = loadedShape.Left;
            int loadedTop = loadedShape.Top;
            int loadedWidth = loadedShape.Width;
            int loadedHeight = loadedShape.Height;

            // -------------------- Validate consistency --------------------
            bool isConsistent = originalLeft == loadedLeft &&
                               originalTop == loadedTop &&
                               originalWidth == loadedWidth &&
                               originalHeight == loadedHeight;

            Console.WriteLine("Slicer placement validation result: " + (isConsistent ? "Consistent" : "Inconsistent"));
            Console.WriteLine($"Original - Left:{originalLeft}, Top:{originalTop}, Width:{originalWidth}, Height:{originalHeight}");
            Console.WriteLine($"Loaded   - Left:{loadedLeft}, Top:{loadedTop}, Width:{loadedWidth}, Height:{loadedHeight}");
        }
    }
}