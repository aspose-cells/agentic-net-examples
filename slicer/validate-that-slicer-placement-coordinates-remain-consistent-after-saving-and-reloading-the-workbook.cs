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
            // ---------- Create workbook and add data ----------
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Sample data for a pivot table
            cells["A1"].Value = "Category";
            cells["B1"].Value = "Amount";
            cells["A2"].Value = "A";
            cells["B2"].Value = 100;
            cells["A3"].Value = "B";
            cells["B3"].Value = 200;
            cells["A4"].Value = "A";
            cells["B4"].Value = 150;
            cells["A5"].Value = "B";
            cells["B5"].Value = 250;

            // ---------- Create a pivot table ----------
            int pivotIdx = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // ---------- Add a slicer linked to the pivot table ----------
            // Place the slicer starting at cell G1
            int slicerIdx = sheet.Slicers.Add(pivot, "G1", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Set explicit placement coordinates using the underlying Shape object
            // (Left and Top are in pixels)
            slicer.Shape.Left = 300;   // horizontal offset from worksheet left border
            slicer.Shape.Top = 100;    // vertical offset from worksheet top border
            slicer.Shape.Width = 150;
            slicer.Shape.Height = 120;

            // Store the coordinates for later comparison
            int originalLeft = slicer.Shape.Left;
            int originalTop = slicer.Shape.Top;

            // ---------- Save the workbook ----------
            string filePath = "SlicerPlacementDemo.xlsx";
            workbook.Save(filePath);

            // ---------- Load the workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath);
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Slicer loadedSlicer = loadedSheet.Slicers[0]; // assume only one slicer

            // Retrieve placement coordinates after reload
            int loadedLeft = loadedSlicer.Shape.Left;
            int loadedTop = loadedSlicer.Shape.Top;

            // ---------- Validate consistency ----------
            bool leftMatches = originalLeft == loadedLeft;
            bool topMatches = originalTop == loadedTop;

            Console.WriteLine($"Original Left: {originalLeft}, Loaded Left: {loadedLeft}, Match: {leftMatches}");
            Console.WriteLine($"Original Top: {originalTop}, Loaded Top: {loadedTop}, Match: {topMatches}");

            if (leftMatches && topMatches)
            {
                Console.WriteLine("Slicer placement coordinates are consistent after save and reload.");
            }
            else
            {
                Console.WriteLine("Slicer placement coordinates changed after save and reload.");
            }
        }
    }
}