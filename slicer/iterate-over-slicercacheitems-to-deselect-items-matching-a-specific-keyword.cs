using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    class DeselectSlicerItemsByKeyword
    {
        static void Main()
        {
            // Load an existing workbook that contains a slicer
            Workbook workbook = new Workbook("input.xlsx");
            // Ensure all worksheets are refreshed (required for slicer data)
            workbook.Worksheets.RefreshAll();

            // Define the keyword for which slicer items should be deselected
            string keyword = "Apple";

            // Iterate through all worksheets to find slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // If the worksheet has no slicers, continue to next sheet
                if (sheet.Slicers.Count == 0) continue;

                // Process each slicer in the worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Access the collection of slicer cache items
                    SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;

                    // Iterate over each cache item
                    foreach (SlicerCacheItem item in cacheItems)
                    {
                        // Deselect the item if its value contains the specified keyword
                        if (item.Value != null && item.Value.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                        {
                            item.Selected = false;
                        }
                    }

                    // Refresh the slicer to apply the selection changes
                    slicer.Refresh();
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}