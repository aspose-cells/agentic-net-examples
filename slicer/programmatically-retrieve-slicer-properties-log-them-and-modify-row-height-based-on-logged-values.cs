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

                // Add a pivot table based on the data
                int pivotIdx = sheet.PivotTables.Add("A1:B4", "D1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Data, "Amount");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table.
                // The second argument must be a valid cell address (e.g., "E1").
                int slicerIdx = sheet.Slicers.Add(pivot, "E1", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];
                slicer.Name = "CategorySlicer";

                // Retrieve and log various slicer properties
                Console.WriteLine("=== Slicer Properties Before Modification ===");
                Console.WriteLine($"Name: {slicer.Name}");
                Console.WriteLine($"Caption: {slicer.Caption}");
                Console.WriteLine($"RowHeight (points): {slicer.RowHeight}");
                Console.WriteLine($"RowHeightPixel: {slicer.RowHeightPixel}");
                Console.WriteLine($"ColumnWidth (points): {slicer.ColumnWidth}");
                Console.WriteLine($"NumberOfColumns: {slicer.NumberOfColumns}");
                Console.WriteLine($"LockedPosition: {slicer.LockedPosition}");

                // Example logic: increase row height by 5 points if current height is less than 25 points
                double currentRowHeight = slicer.RowHeight;
                if (currentRowHeight < 25)
                {
                    double newRowHeight = currentRowHeight + 5;
                    slicer.RowHeight = newRowHeight;
                    Console.WriteLine($"RowHeight increased from {currentRowHeight} to {newRowHeight} points.");
                }
                else
                {
                    Console.WriteLine("RowHeight is already 25 points or more; no change applied.");
                }

                // Log properties after modification
                Console.WriteLine("=== Slicer Properties After Modification ===");
                Console.WriteLine($"RowHeight (points): {slicer.RowHeight}");

                // Save the workbook
                string outputPath = "SlicerPropertyLogAndModify.xlsx";
                // Ensure the directory exists (use current directory if none specified)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}