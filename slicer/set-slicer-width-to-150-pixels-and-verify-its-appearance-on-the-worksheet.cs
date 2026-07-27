using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerWidthDemo
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

                // Populate worksheet with sample data for a pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["A3"].PutValue("Fruit");
                sheet.Cells["A4"].PutValue("Vegetable");
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["B4"].PutValue(150);

                // Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");

                // Add a slicer linked to the pivot table (field name "Category")
                int slicerIdx = sheet.Slicers.Add(pivot, "Category", "F1");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Set the slicer width (using Shape.Width, measured in points)
                slicer.Shape.Width = 150; // Approx. 150 points

                // Verify the width by reading the property and printing to console
                Console.WriteLine($"Slicer width set to {slicer.Shape.Width} points.");

                // Save the workbook to file
                string outputPath = "SlicerWidthDemo.xlsx";
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