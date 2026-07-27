using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerResizeDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["A2"].Value = "Fruit";
                sheet.Cells["A3"].Value = "Fruit";
                sheet.Cells["A4"].Value = "Vegetable";

                sheet.Cells["B1"].Value = "Sales";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["B3"].Value = 150;
                sheet.Cells["B4"].Value = 200;

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

                // Add a slicer linked to the pivot table (use field name instead of index)
                int slicerIdx = sheet.Slicers.Add(pivot, "Category", "F1");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Resize the slicer using the Shape properties (points)
                slicer.Shape.WidthPt = 200;   // Width in points
                slicer.Shape.HeightPt = 120;  // Height in points

                // Optionally lock the slicer position to prevent accidental moves
                slicer.LockedPosition = true;

                // Save the workbook
                string outputPath = "SlicerResized.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}