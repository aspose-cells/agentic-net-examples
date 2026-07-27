using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerCustomColorDemo
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Define a custom color palette (modify index 55)
            // -------------------------------------------------
            // Change palette entry 55 to a custom light orange color
            Color customColor = Color.FromArgb(255, 230, 180);
            workbook.ChangePalette(customColor, 55);

            // -------------------------------------------------
            // 2. Prepare sample data on the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            // Header
            cells["A1"].PutValue("Fruit");
            cells["B1"].PutValue("Quantity");

            // Data rows
            cells["A2"].PutValue("Apple");
            cells["B2"].PutValue(10);
            cells["A3"].PutValue("Orange");
            cells["B3"].PutValue(15);
            cells["A4"].PutValue("Banana");
            cells["B4"].PutValue(20);

            // -------------------------------------------------
            // 3. Create a pivot table based on the data
            // -------------------------------------------------
            int pivotIdx = dataSheet.PivotTables.Add("A1:B4", "D3", "FruitPivot");
            PivotTable pivot = dataSheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, 0);      // Fruit column
            pivot.AddFieldToArea(PivotFieldType.Data, 1);     // Quantity column
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // 4. Add a slicer on the first worksheet and apply a built‑in style
            // -------------------------------------------------
            int slicerIdx1 = dataSheet.Slicers.Add(pivot, "F3", "Fruit");
            Slicer slicer1 = dataSheet.Slicers[slicerIdx1];
            slicer1.StyleType = SlicerStyleType.SlicerStyleDark1; // uses palette colors
            slicer1.Caption = "Fruit Selector";
            slicer1.NumberOfColumns = 1;
            slicer1.WidthPixel = 150;
            slicer1.HeightPixel = 100;

            // -------------------------------------------------
            // 5. Add a second worksheet and place an identical slicer
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            // Add the same slicer connected to the same pivot table
            int slicerIdx2 = sheet2.Slicers.Add(pivot, "A1", "Fruit");
            Slicer slicer2 = sheet2.Slicers[slicerIdx2];
            slicer2.StyleType = SlicerStyleType.SlicerStyleDark1; // same style ensures visual consistency
            slicer2.Caption = "Fruit Selector (Sheet2)";
            slicer2.NumberOfColumns = 1;
            slicer2.WidthPixel = 150;
            slicer2.HeightPixel = 100;

            // -------------------------------------------------
            // 6. Refresh slicers to reflect any data changes
            // -------------------------------------------------
            slicer1.Refresh();
            slicer2.Refresh();

            // -------------------------------------------------
            // 7. Save the workbook
            // -------------------------------------------------
            workbook.Save("SlicerCustomColorConsistency.xlsx");
        }
    }
}