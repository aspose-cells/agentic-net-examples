using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook that contains slicers
            Workbook workbook = new Workbook("InputWithSlicers.xlsx");

            // Assume the slicer is on the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the slicer collection for the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Check if there is at least one slicer to remove
            if (slicers.Count > 0)
            {
                // Option 1: Remove by index (e.g., the first slicer)
                slicers.RemoveAt(0);

                // Option 2: Remove by reference
                // Slicer slicerToRemove = slicers[0];
                // slicers.Remove(slicerToRemove);
            }

            // Save the workbook after removal
            workbook.Save("OutputWithoutSlicer.xlsx");
        }
    }
}