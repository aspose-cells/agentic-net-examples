using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

namespace AsposeCellsSlicerBorderDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pivot table
                sheet.Cells["A1"].Value = "Fruit";
                sheet.Cells["A2"].Value = "Apple";
                sheet.Cells["A3"].Value = "Orange";
                sheet.Cells["A4"].Value = "Banana";

                sheet.Cells["B1"].Value = "Quantity";
                sheet.Cells["B2"].Value = 10;
                sheet.Cells["B3"].Value = 15;
                sheet.Cells["B4"].Value = 20;

                // Add a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "E3", "FruitPivot");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                pivot.AddFieldToArea(PivotFieldType.Data, 1);

                // Add a slicer linked to the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "G1", pivot.BaseFields[0]);
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Access the underlying shape of the slicer
                Shape slicerShape = slicer.Shape;

                // Ensure the slicer has a visible line (border)
                slicerShape.HasLine = true;

                // Apply custom border: dashed line with defined thickness
                slicerShape.Line.Weight = 2.0f;                           // Thickness (points)
                slicerShape.Line.DashStyle = MsoLineDashStyle.Dash;       // Dashed style

                // Optional: adjust slicer size and caption
                slicer.Caption = "Fruit Selection";
                slicer.Shape.WidthPt = 200;
                slicer.Shape.HeightPt = 100;

                // Save the workbook
                string outputPath = "SlicerCustomBorderDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}