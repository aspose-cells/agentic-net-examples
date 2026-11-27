// Title: How to set a slicer width to 150 pixels and read its properties with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set slicer width in pixels | Read slicer WidthPixel value after assignment in Aspose.Cells | Programmatically adjust slicer dimensions for a pivot table using Aspose.Cells | Verify slicer size on a worksheet with Aspose.Cells .NET example | Position and resize a pivot table slicer in C# Aspose.Cells
// Tags: Aspose.Cells slicer width configuration | C# Aspose.Cells slicer resizing | Aspose.Cells pivot slicer layout | Aspose.Cells retrieve slicer metrics | Aspose.Cells export workbook with slicer

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace SlicerWidthDemo
{
    // The example creates a workbook, builds a pivot table from sample data, adds a slicer linked to the 'Category' field, positions it, sets its WidthPixel to 150 (HeightPixel to 200), prints the slicer's WidthPixel, HeightPixel, TopPixel, and LeftPixel to the console, and saves the file as SlicerWidthDemo.xlsx for visual verification.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data for a pivot table
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Product");
                sheet.Cells["C1"].PutValue("Quantity");

                sheet.Cells["A2"].PutValue("Fruits");
                sheet.Cells["B2"].PutValue("Apple");
                sheet.Cells["C2"].PutValue(10);

                sheet.Cells["A3"].PutValue("Fruits");
                sheet.Cells["B3"].PutValue("Banana");
                sheet.Cells["C3"].PutValue(15);

                sheet.Cells["A4"].PutValue("Vegetables");
                sheet.Cells["B4"].PutValue("Carrot");
                sheet.Cells["C4"].PutValue(8);

                // 3. Add a pivot table based on the data range
                int pivotIdx = sheet.PivotTables.Add("A1:C4", "E1", "PivotTable1");
                PivotTable pivot = sheet.PivotTables[pivotIdx];

                // Row field – Category
                pivot.AddFieldToArea(PivotFieldType.Row, 0);
                // Column field – Product
                pivot.AddFieldToArea(PivotFieldType.Column, 1);
                // Data field – Quantity (sum)
                pivot.AddFieldToArea(PivotFieldType.Data, 2);
                pivot.RefreshData();
                pivot.CalculateData();

                // 4. Add a slicer linked to the "Category" field of the pivot table
                int slicerIdx = sheet.Slicers.Add(pivot, "Category", "CategorySlicer");
                Slicer slicer = sheet.Slicers[slicerIdx];

                // Position the slicer (optional, for better visibility)
                slicer.TopPixel = 100;
                slicer.LeftPixel = 400;

                // 5. Set the slicer width to 150 pixels
                slicer.WidthPixel = 150; // Exact API to set width in pixels

                // Optional: set a height to see the slicer clearly
                slicer.HeightPixel = 200;

                // 6. Verify the width by reading back the property
                Console.WriteLine("Slicer \"CategorySlicer\" properties:");
                Console.WriteLine($"  Width (pixels) : {slicer.WidthPixel}");
                Console.WriteLine($"  Height (pixels): {slicer.HeightPixel}");
                Console.WriteLine($"  Top (pixels)   : {slicer.TopPixel}");
                Console.WriteLine($"  Left (pixels)  : {slicer.LeftPixel}");

                // 7. Save the workbook to a file
                string outputPath = "SlicerWidthDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\". Open the file to visually confirm the slicer width.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the slicer:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
