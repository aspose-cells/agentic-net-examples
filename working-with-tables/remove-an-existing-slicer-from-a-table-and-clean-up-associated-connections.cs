using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class RemoveSlicerExample
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Assume the slicer is on the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the slicer collection of the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Proceed only if there is at least one slicer
        if (slicers.Count > 0)
        {
            // Retrieve the slicer to be removed (e.g., the first one)
            Slicer slicerToRemove = slicers[0];

            // Clean up any PivotTable connections the slicer may have
            foreach (PivotTable pivot in worksheet.PivotTables)
            {
                slicerToRemove.RemovePivotConnection(pivot);
            }

            // Remove the slicer from the collection
            slicers.Remove(slicerToRemove);
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}