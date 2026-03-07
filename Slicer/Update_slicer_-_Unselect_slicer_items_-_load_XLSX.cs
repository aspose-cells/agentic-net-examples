using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerUnselectDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Process each slicer in the collection
                for (int s = 0; s < slicers.Count; s++)
                {
                    Slicer slicer = slicers[s];

                    // Unselect every slicer item
                    for (int i = 0; i < slicer.SlicerCache.SlicerCacheItems.Count; i++)
                    {
                        slicer.SlicerCache.SlicerCacheItems[i].Selected = false;
                    }

                    // Refresh the slicer so the changes take effect
                    slicer.Refresh();
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}