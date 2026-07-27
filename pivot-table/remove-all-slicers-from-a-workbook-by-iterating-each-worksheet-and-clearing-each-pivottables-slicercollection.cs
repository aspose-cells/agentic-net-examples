using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class RemoveAllSlicers
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Clear all slicers from this worksheet
            slicers.Clear();
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}