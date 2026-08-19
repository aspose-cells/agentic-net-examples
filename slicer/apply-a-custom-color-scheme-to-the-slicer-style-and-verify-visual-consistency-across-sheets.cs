// Title: Aspose.Cells .NET – Apply a Custom Color Palette to Slicer Styles for Consistent Appearance Across Worksheets
// Description: Shows how to build a workbook, add sample data and a pivot table, change the workbook palette with a custom color at index 55, place slicers on two worksheets using a built‑in style that picks the palette entry, refresh the slicers, and save the file so the custom color renders uniformly on all slicers.
// Keywords: Aspose.Cells | C# | .NET | slicer | custom color palette | workbook palette | SlicerStyleDark2 | pivot table | multiple worksheets | visual consistency | Excel automation
// Common Searches: Aspose.Cells change slicer color palette .NET | apply custom color to slicer style across sheets | slicer style consistency with workbook palette | C# Aspose.Cells custom slicer color example | verify slicer visual appearance after palette change
// Developer Intent: Create a .NET workbook where slicers on different sheets share the same custom color defined via the workbook palette.
// Use Cases: Add a light‑orange color to palette index 55 and use SlicerStyleDark2 so slicers on Sheet1 and Sheet2 display the same hue. | Refresh slicers after modifying the palette to ensure the new color is applied instantly. | Duplicate pivot tables on additional worksheets while preserving the custom slicer styling.
// AI Prompts: Generate C# code that sets a teal color at palette index 55 and applies a built‑in slicer style that uses this entry on several worksheets. | Explain how Aspose.Cells maps palette entries to slicer style colors and outline steps to confirm uniform slicer appearance across sheets. | Write unit tests in C# with Aspose.Cells to assert that slicers on different worksheets have identical background colors after a palette update.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerCustomColorSchemeDemo
{
    // Shows how to build a workbook, add sample data and a pivot table, change the workbook palette with a custom color at index 55, place slicers on two worksheets using a built‑in style that picks the palette entry, refresh the slicers, and save the file so the custom color renders uniformly on all slicers.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare sample data for the pivot table
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            Cells cells = dataSheet.Cells;

            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Item");
            cells["C1"].PutValue("Amount");

            cells["A2"].PutValue("Fruit");
            cells["B2"].PutValue("Apple");
            cells["C2"].PutValue(120);

            cells["A3"].PutValue("Fruit");
            cells["B3"].PutValue("Banana");
            cells["C3"].PutValue(80);

            cells["A4"].PutValue("Vegetable");
            cells["B4"].PutValue("Carrot");
            cells["C4"].PutValue(150);

            cells["A5"].PutValue("Vegetable");
            cells["B5"].PutValue("Tomato");
            cells["C5"].PutValue(90);

            // -------------------------------------------------
            // 2. Create a pivot table based on the data
            // -------------------------------------------------
            int pivotIndex = dataSheet.PivotTables.Add("A1:C5", "E1", "SalesPivot");
            PivotTable pivot = dataSheet.PivotTables[pivotIndex];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Column, "Item");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.RefreshData();
            pivot.CalculateData();

            // -------------------------------------------------
            // 3. Change the workbook palette to introduce a custom color
            //    (index 55 is the last palette entry)
            // -------------------------------------------------
            Color customColor = Color.FromArgb(255, 200, 150); // Light orange
            workbook.ChangePalette(customColor, 55);

            // -------------------------------------------------
            // 4. Add a slicer on the first sheet and apply a built‑in style
            //    that will use the modified palette entry.
            // -------------------------------------------------
            SlicerCollection slicersSheet1 = dataSheet.Slicers;
            int slicerIdx1 = slicersSheet1.Add(pivot, "G2", "Category");
            Slicer slicer1 = slicersSheet1[slicerIdx1];
            slicer1.StyleType = SlicerStyleType.SlicerStyleDark2; // Built‑in style
            slicer1.NumberOfColumns = 2;
            slicer1.Caption = "Category Filter (Sheet1)";

            // -------------------------------------------------
            // 5. Add a second worksheet and place an identical slicer there
            // -------------------------------------------------
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            // Copy the same data to keep the pivot reference valid (optional)
            sheet2.Cells.CopyRows(dataSheet.Cells, 0, 0, dataSheet.Cells.MaxDataRow + 1);
            // Add the same pivot table on the second sheet (required for slicer connection)
            int pivotIdx2 = sheet2.PivotTables.Add("A1:C5", "E1", "SalesPivot2");
            PivotTable pivot2 = sheet2.PivotTables[pivotIdx2];
            pivot2.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot2.AddFieldToArea(PivotFieldType.Column, "Item");
            pivot2.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot2.RefreshData();
            pivot2.CalculateData();

            // Add slicer on the second sheet
            SlicerCollection slicersSheet2 = sheet2.Slicers;
            int slicerIdx2 = slicersSheet2.Add(pivot2, "G2", "Category");
            Slicer slicer2 = slicersSheet2[slicerIdx2];
            slicer2.StyleType = SlicerStyleType.SlicerStyleDark2; // Same style as slicer1
            slicer2.NumberOfColumns = 2;
            slicer2.Caption = "Category Filter (Sheet2)";

            // -------------------------------------------------
            // 6. Refresh slicers to ensure they reflect the latest data
            // -------------------------------------------------
            slicer1.Refresh();
            slicer2.Refresh();

            // -------------------------------------------------
            // 7. Save the workbook
            // -------------------------------------------------
            workbook.Save("SlicerCustomColorSchemeDemo.xlsx");
        }
    }
}
