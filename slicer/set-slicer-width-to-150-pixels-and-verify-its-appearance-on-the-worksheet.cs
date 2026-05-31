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
                sheet.Cells["B1"].PutValue("Amount");
                sheet.Cells["A2"].PutValue("Fruit");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["A3"].PutValue("Vegetable");
                sheet.Cells["B3"].PutValue(80);
                sheet.Cells["A4"].PutValue("Fruit");
                sheet.Cells["B4"].PutValue(150);
                sheet.Cells["A5"].PutValue("Vegetable");
                sheet.Cells["B5"].PutValue(70);

                // Add a pivot table based on the data range
                PivotTableCollection pivots = sheet.PivotTables;
                int pivotIdx = pivots.Add("A1:B5", "D2", "PivotTable1");
                PivotTable pivot = pivots[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table for the "Category" field
                // Correct parameter order: destination cell first, then the field name
                int slicerIdx = sheet.Slicers.Add(pivot, "F2", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Set the slicer width to 150 pixels
                slicer.WidthPixel = 150;

                // Verify that the width was set correctly
                if (slicer.WidthPixel == 150)
                {
                    Console.WriteLine("Slicer width successfully set to 150 pixels.");
                }
                else
                {
                    Console.WriteLine($"Unexpected slicer width: {slicer.WidthPixel} pixels.");
                }

                // Save the workbook to a file
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