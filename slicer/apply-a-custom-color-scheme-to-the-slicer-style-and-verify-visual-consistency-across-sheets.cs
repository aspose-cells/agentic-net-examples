using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerCustomColorSchemeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Fruit");
            dataSheet.Cells["A2"].PutValue("Apple");
            dataSheet.Cells["A3"].PutValue("Orange");
            dataSheet.Cells["A4"].PutValue("Banana");
            dataSheet.Cells["B1"].PutValue("Quantity");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["B3"].PutValue(15);
            dataSheet.Cells["B4"].PutValue(20);

            // Add a pivot table on the same sheet
            int pivotIdx = dataSheet.PivotTables.Add("A1:B4", "D3", "FruitPivot");
            PivotTable pivot = dataSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);   // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);  // Quantity column
            pivot.RefreshData();
            pivot.CalculateData();

            // Change the workbook palette to introduce custom colors
            // Index 0-5 correspond to the first six palette entries
            workbook.ChangePalette(Color.FromArgb(255, 200, 200), 0); // Light red
            workbook.ChangePalette(Color.FromArgb(200, 255, 200), 1); // Light green
            workbook.ChangePalette(Color.FromArgb(200, 200, 255), 2); // Light blue
            workbook.ChangePalette(Color.FromArgb(255, 255, 200), 3); // Light yellow
            workbook.ChangePalette(Color.FromArgb(255, 200, 255), 4); // Light magenta
            workbook.ChangePalette(Color.FromArgb(200, 255, 255), 5); // Light cyan

            // Add a slicer on the first sheet and apply a built‑in style
            int slicerIdx1 = dataSheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer1 = dataSheet.Slicers[slicerIdx1];
            slicer1.StyleType = SlicerStyleType.SlicerStyleDark2; // Uses palette colors
            slicer1.Caption = "Fruit Selection";
            slicer1.NumberOfColumns = 2;
            slicer1.WidthPixel = 250;
            slicer1.HeightPixel = 120;

            // Add a second worksheet to verify visual consistency across sheets
            Worksheet secondSheet = workbook.Worksheets.Add("Verification");
            // Add a slicer on the second sheet that connects to the same pivot table
            int slicerIdx2 = secondSheet.Slicers.Add(pivot, "A1", "Fruit");
            Slicer slicer2 = secondSheet.Slicers[slicerIdx2];
            slicer2.StyleType = SlicerStyleType.SlicerStyleDark2; // Same style as first slicer
            slicer2.Caption = "Fruit Selection (Copy)";
            slicer2.NumberOfColumns = 2;
            slicer2.WidthPixel = 250;
            slicer2.HeightPixel = 120;

            // Refresh slicers to ensure they reflect the latest data and style
            slicer1.Refresh();
            slicer2.Refresh();

            // Save the workbook
            workbook.Save("SlicerCustomColorSchemeDemo.xlsx");
        }
    }
}