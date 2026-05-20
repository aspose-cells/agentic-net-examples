using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class UpdateSlicerTitle
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // If there is at least one slicer, change its title
        if (slicers.Count > 0)
        {
            Slicer slicer = slicers[0];
            slicer.Title = "Region Filter"; // Set the new title
        }

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}