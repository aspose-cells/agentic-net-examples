using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class RemoveSlicerDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // The name of the slicer that should be removed
        string slicerNameToRemove = "FruitSlicer1";

        // Search each worksheet for the slicer with the specified name
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            SlicerCollection slicers = sheet.Slicers;

            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Slicer.Name holds the slicer's name
                if (slicer.Name == slicerNameToRemove)
                {
                    // Remove the slicer from the collection
                    slicers.Remove(slicer);
                    // Exit the loops after removal
                    break;
                }
            }
        }

        // Save the modified workbook as XLSX
        workbook.Save("output.xlsx");
    }
}