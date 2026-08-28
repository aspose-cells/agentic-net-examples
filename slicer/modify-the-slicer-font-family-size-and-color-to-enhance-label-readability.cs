// Title: Set slicer label font family, size, and color in an Aspose.Cells .NET workbook using C#
// AI Prompts: Create a workbook with a pivot table, add a slicer for the "Category" field, and programmatically set the slicer’s label font to Calibri 12 pt dark blue using Aspose.Cells in C#. | Update the font properties (Name, Size, Color) of an existing slicer linked to a pivot table in an Aspose.Cells workbook via the slicer’s Shape.Font object. | Apply a built‑in slicer style and then override its label font settings (family, size, color) in a .NET spreadsheet generated with Aspose.Cells.
// Common Searches: Aspose.Cells C# change slicer label font family and size | How to set slicer text color in a .NET workbook with Aspose.Cells | Programmatically customize slicer appearance in Excel using Aspose.Cells for .NET | C# code to modify slicer font properties after creating a pivot table with Aspose.Cells
// Tags: Aspose.Cells slicer text styling C# | modify slicer font settings .NET | apply Shape.Font to slicer for styling | pivot table slicer visual tweaks | programmatic slicer label color change

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerFontCustomization
{
    // The example creates a workbook, fills it with sample data, builds a pivot table, adds a slicer linked to the "Category" field, applies a built‑in slicer style, and then customizes the slicer’s label font to Calibri 12 pt dark blue via the Shape.Font properties before saving the file as SlicerFontCustomization.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].Value = "Category";
            sheet.Cells["A2"].Value = "Fruit";
            sheet.Cells["A3"].Value = "Fruit";
            sheet.Cells["A4"].Value = "Vegetable";
            sheet.Cells["A5"].Value = "Vegetable";

            sheet.Cells["B1"].Value = "Item";
            sheet.Cells["B2"].Value = "Apple";
            sheet.Cells["B3"].Value = "Banana";
            sheet.Cells["B4"].Value = "Carrot";
            sheet.Cells["B5"].Value = "Potato";

            sheet.Cells["C1"].Value = "Quantity";
            sheet.Cells["C2"].Value = 10;
            sheet.Cells["C3"].Value = 15;
            sheet.Cells["C4"].Value = 20;
            sheet.Cells["C5"].Value = 25;

            // Create a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:C5", "E1", "PivotTable1");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Category");
            pivot.AddFieldToArea(PivotFieldType.Row, "Item");
            pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the "Category" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "G1", "Category");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Optional: set a built‑in style for the slicer
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

            // Access the underlying shape of the slicer and modify its font
            // Font family (Name), size, and color are set to improve readability
            slicer.Shape.Font.Name = "Calibri";
            slicer.Shape.Font.Size = 12;               // Font size in points
            slicer.Shape.Font.Color = Color.DarkBlue; // Font color

            // Save the workbook with the customized slicer
            workbook.Save("SlicerFontCustomization.xlsx");
        }
    }
}
