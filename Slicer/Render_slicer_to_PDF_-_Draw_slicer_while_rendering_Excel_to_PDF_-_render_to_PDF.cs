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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the pivot table
            cells["A1"].Value = "Fruit";
            cells["B1"].Value = "Year";
            cells["C1"].Value = "Amount";

            cells["A2"].Value = "Apple";
            cells["B2"].Value = 2020;
            cells["C2"].Value = 50;

            cells["A3"].Value = "Apple";
            cells["B3"].Value = 2021;
            cells["C3"].Value = 70;

            cells["A4"].Value = "Banana";
            cells["B4"].Value = 2020;
            cells["C4"].Value = 30;

            cells["A5"].Value = "Banana";
            cells["B5"].Value = 2021;
            cells["C5"].Value = 45;

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:C5", "E1", "FruitPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];
            pivot.AddFieldToArea(PivotFieldType.Row, "Fruit");
            pivot.AddFieldToArea(PivotFieldType.Column, "Year");
            pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
            pivot.PivotTableStyleType = PivotTableStyleType.PivotTableStyleMedium9;
            pivot.RefreshData();
            pivot.CalculateData();

            // Add a slicer for the "Fruit" field of the pivot table
            int slicerIdx = sheet.Slicers.Add(pivot, "G1", "Fruit");
            Slicer slicer = sheet.Slicers[slicerIdx];

            // Make sure the slicer is printable (visible in PDF)
            // Use the Shape's IsPrintable property (recommended)
            slicer.Shape.IsPrintable = true;

            // Optionally set some visual properties
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2;
            slicer.Caption = "Select Fruit";

            // Save the workbook as PDF – slicer will be rendered because it is printable
            workbook.Save("SlicerDemo.pdf", SaveFormat.Pdf);

            Console.WriteLine("Workbook with slicer saved to SlicerDemo.pdf");
        }
    }
}