using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Loop through each slicer on the current worksheet
            foreach (Slicer slicer in sheet.Slicers)
            {
                // Deselect every slicer item
                foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
                {
                    item.Selected = false;
                }

                // Refresh the slicer (also refreshes its PivotTable)
                slicer.Refresh();
            }
        }

        // Save the workbook with the updated slicer state
        workbook.Save("output.xlsx");
    }
}