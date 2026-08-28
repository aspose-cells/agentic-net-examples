// Title: How to read slicer properties and programmatically increase its row height in Aspose.Cells C# pivot tables
// AI Prompts: Write C# code using Aspose.Cells to enumerate all properties of a slicer attached to a pivot table and output them to the console. | Create a conditional statement that checks a slicer's RowHeight and, if it is less than 25 points, adds 5 points and updates RowHeightPixel accordingly. | Save the modified workbook to a new file while preserving the original pivot table and slicer configuration.
// Common Searches: Aspose.Cells C# retrieve slicer RowHeight and ColumnWidth values | increase slicer row height programmatically in Aspose.Cells .NET | log slicer properties before modifying them with Aspose.Cells | convert slicer row height from points to pixels using Aspose.Cells | adjust slicer dimensions after creating pivot table in C#
// Tags: Aspose.Cells slicer property extraction C# | Aspose.Cells adjust slicer row height | Aspose.Cells pivot table slicer dimensions | Aspose.Cells points to pixel conversion for slicer | Aspose.Cells save workbook after slicer changes

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerPropertyDemo
{
    // The example creates a workbook, adds sample data, builds a pivot table, inserts a slicer linked to the Category field, logs various slicer properties (name, caption, row height, pixel height, column width, etc.), increases the slicer row height by 5 points when it is under 25 points, updates the corresponding pixel height, and saves the workbook as SlicerPropertyModified.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate worksheet with sample data for a pivot table
                sheet.Cells["A1"].Value = "Category";
                sheet.Cells["B1"].Value = "Product";
                sheet.Cells["C1"].Value = "Quantity";

                sheet.Cells["A2"].Value = "Fruit";
                sheet.Cells["B2"].Value = "Apple";
                sheet.Cells["C2"].Value = 50;

                sheet.Cells["A3"].Value = "Fruit";
                sheet.Cells["B3"].Value = "Banana";
                sheet.Cells["C3"].Value = 30;

                sheet.Cells["A4"].Value = "Vegetable";
                sheet.Cells["B4"].Value = "Carrot";
                sheet.Cells["C4"].Value = 20;

                // Add a pivot table based on the sample data
                int pivotIdx = sheet.PivotTables.Add("A1:C4", "E5", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];
                pivot.AddFieldToArea(PivotFieldType.Row, "Category");
                pivot.AddFieldToArea(PivotFieldType.Row, "Product");
                pivot.AddFieldToArea(PivotFieldType.Data, "Quantity");
                pivot.RefreshData();
                pivot.CalculateData();

                // Add a slicer linked to the pivot table (filtering by Category)
                // Destination cell for the slicer (top‑left corner) is "G5"
                int slicerIdx = sheet.Slicers.Add(pivot, "G5", "Category");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Retrieve and log slicer properties
                Console.WriteLine("=== Slicer Properties Before Modification ===");
                Console.WriteLine($"Name: {slicer.Name}");
                Console.WriteLine($"Caption: {slicer.Caption}");
                Console.WriteLine($"RowHeight (points): {slicer.RowHeight}");
                Console.WriteLine($"RowHeightPixel (pixels): {slicer.RowHeightPixel}");
                Console.WriteLine($"ColumnWidth (points): {slicer.ColumnWidth}");
                Console.WriteLine($"NumberOfColumns: {slicer.NumberOfColumns}");
                Console.WriteLine($"Height (points) [obsolete]: {slicer.Height}");
                Console.WriteLine($"Width (points) [obsolete]: {slicer.Width}");
                Console.WriteLine($"LockedPosition: {slicer.LockedPosition}");
                Console.WriteLine($"StyleType: {slicer.StyleType}");

                // Modify the row height based on the logged RowHeight value
                // Example rule: increase each row height by 5 points if current height is less than 25 points
                double currentRowHeight = slicer.RowHeight;
                if (currentRowHeight < 25)
                {
                    slicer.RowHeight = currentRowHeight + 5;
                    Console.WriteLine($"RowHeight increased from {currentRowHeight} to {slicer.RowHeight}");
                }
                else
                {
                    Console.WriteLine("RowHeight is already 25 points or more; no change applied.");
                }

                // Adjust RowHeightPixel to keep pixel/point consistency (1 point ≈ 1.33333 pixels)
                slicer.RowHeightPixel = (int)Math.Round(slicer.RowHeight * 1.33333);
                Console.WriteLine($"RowHeightPixel adjusted to {slicer.RowHeightPixel}");

                // Save the workbook with the modified slicer
                workbook.Save("SlicerPropertyModified.xlsx");
                Console.WriteLine("Workbook saved as SlicerPropertyModified.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
