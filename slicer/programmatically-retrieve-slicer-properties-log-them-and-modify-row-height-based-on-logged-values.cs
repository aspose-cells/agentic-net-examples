using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerPropertyDemo
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
                sheet.Cells["B1"].Value = "Amount";
                sheet.Cells["A2"].Value = "Fruit";
                sheet.Cells["B2"].Value = 120;
                sheet.Cells["A3"].Value = "Fruit";
                sheet.Cells["B3"].Value = 80;
                sheet.Cells["A4"].Value = "Vegetable";
                sheet.Cells["B4"].Value = 150;
                sheet.Cells["A5"].Value = "Vegetable";
                sheet.Cells["B5"].Value = 70;

                // Add a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B5", "D3", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table.
                // Destination cell must be a valid address (e.g., "E3").
                int slicerIdx = sheet.Slicers.Add(pivot, "E3", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Retrieve and log slicer properties
                Console.WriteLine("=== Slicer Initial Properties ===");
                Console.WriteLine($"Name: {slicer.Name}");
                Console.WriteLine($"Caption: {slicer.Caption}");
                Console.WriteLine($"RowHeight (points): {slicer.RowHeight}");
                Console.WriteLine($"RowHeightPixel: {slicer.RowHeightPixel}");
                Console.WriteLine($"ColumnWidth (points): {slicer.ColumnWidth}");
                Console.WriteLine($"NumberOfColumns: {slicer.NumberOfColumns}");
                Console.WriteLine($"LockedPosition: {slicer.LockedPosition}");
                Console.WriteLine($"StyleType: {slicer.StyleType}");

                // Example logic: if the current RowHeight is less than 25 points, increase it to 30 points
                double currentRowHeight = slicer.RowHeight;
                if (currentRowHeight < 25)
                {
                    double newHeight = 30;
                    slicer.RowHeight = newHeight;
                    Console.WriteLine($"RowHeight was {currentRowHeight} pts, updated to {newHeight} pts.");
                }
                else
                {
                    Console.WriteLine($"RowHeight ({currentRowHeight} pts) meets the required minimum; no change applied.");
                }

                // Save the workbook
                string outputPath = "SlicerPropertyDemo.xlsx";
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