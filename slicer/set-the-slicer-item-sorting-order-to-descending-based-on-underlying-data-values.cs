using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class SetSlicerSortOrder
{
    static void Main()
    {
        // Load an existing workbook that contains a slicer
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the slicer collection from the worksheet
        SlicerCollection slicers = sheet.Slicers;

        // Ensure there is at least one slicer present
        if (slicers.Count > 0)
        {
            // Get the first slicer in the collection
            Slicer slicer = slicers[0];

            // Set the slicer items to be sorted in descending order
            slicer.SortOrderType = SortOrder.Descending;
        }

        // Save the workbook with the updated slicer sorting
        workbook.Save("output.xlsx");
    }
}