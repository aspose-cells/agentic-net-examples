using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerProcessingDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains slicers
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Iterate through each slicer in the collection
                for (int s = 0; s < slicers.Count; s++)
                {
                    Slicer slicer = slicers[s];

                    // Access the slicer cache items
                    SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

                    // Log currently selected items
                    Console.WriteLine($"Worksheet: {sheet.Name}, Slicer: {slicer.Name}");
                    for (int i = 0; i < items.Count; i++)
                    {
                        SlicerCacheItem item = items[i];
                        if (item.Selected)
                        {
                            Console.WriteLine($"  Selected Item: {item.Value}");
                        }
                    }

                    // Deselect all items
                    for (int i = 0; i < items.Count; i++)
                    {
                        items[i].Selected = false;
                    }

                    // Refresh the slicer (also refreshes the underlying PivotTable)
                    slicer.Refresh();
                }
            }

            // Save the workbook after modifications
            workbook.Save("output.xlsx");
        }
    }
}