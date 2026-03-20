using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerAdjustmentDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the existing workbook that contains at least one slicer
            string inputPath = "InputWorkbook.xlsx";

            // Load the workbook using the constructor that accepts a file name (lifecycle rule)
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Access the slicer collection
            SlicerCollection slicers = sheet.Slicers;

            // Ensure there is at least one slicer to modify
            if (slicers.Count > 0)
            {
                // Get the first slicer (index 0)
                Slicer slicer = slicers[0];

                // Adjust various slicer properties
                slicer.Caption = "Updated Slicer Caption";
                slicer.ShowCaption = true;                     // Show the header
                slicer.NumberOfColumns = 2;                    // Display items in two columns
                slicer.WidthPixel = 250;                       // Set width in pixels
                slicer.HeightPixel = 180;                      // Set height in pixels
                slicer.LockedPosition = true;                  // Prevent moving/resizing by UI
                slicer.StyleType = SlicerStyleType.SlicerStyleLight3; // Apply a built‑in style
                slicer.ShowAllItems = false;                  // Hide items with no data
                slicer.FirstItemIndex = 0;                     // Start from the first item

                // Refresh the slicer to apply changes and recalculate any linked PivotTables
                slicer.Refresh();
            }
            else
            {
                Console.WriteLine("No slicers found in the worksheet.");
            }

            // Save the modified workbook using the Save method (lifecycle rule)
            string outputPath = "AdjustedWorkbook.xlsx";
            workbook.Save(outputPath);

            // Clean up
            workbook.Dispose();

            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}