using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the slicer (adjust index as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = sheet.Slicers;

        // Ensure there is at least one slicer present
        if (slicers.Count > 0)
        {
            // Retrieve the first slicer (or use a specific name/index)
            Slicer slicer = slicers[0];

            // Programmatically select slicer items.
            // Example: select the first two items and deselect the rest.
            for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
            {
                SlicerCacheItem item = slicer.SlicerCache.SlicerCacheItems[i];
                item.Selected = i < 2; // true for first two items, false otherwise
            }

            // Refresh the slicer (also refreshes associated PivotTables)
            slicer.Refresh();
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}