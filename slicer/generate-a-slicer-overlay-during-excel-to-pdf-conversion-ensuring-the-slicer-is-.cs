using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerPdfDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["A2"].Value = "Apple";
            cells["A3"].Value = "Orange";
            cells["A4"].Value = "Banana";
            cells["B1"].Value = "Sales";
            cells["B2"].Value = 120;
            cells["B3"].Value = 150;
            cells["B4"].Value = 200;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer linked to the pivot table's "Fruit" field
            int slicerIdx = sheet.Slicers.Add(pivot, "E12", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Ensure the slicer is printable so it appears in the PDF output
            // (IsPrintable property is obsolete; the underlying Shape property is preferred)
            slicer.Shape.IsPrintable = true;

            // Optional: adjust slicer appearance
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.WidthPixel = 150;
            slicer.HeightPixel = 200;

            // Save the workbook as PDF; slicer will be rendered in the PDF
            workbook.Save("SlicerOverlay.pdf", SaveFormat.Pdf);
        }
    }
}