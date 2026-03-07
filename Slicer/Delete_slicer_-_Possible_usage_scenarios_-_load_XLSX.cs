using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class DeleteSlicer
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index or name as needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the collection of slicers on this worksheet
        SlicerCollection slicers = sheet.Slicers;

        // Ensure there is at least one slicer to delete
        if (slicers.Count > 0)
        {
            // Delete the first slicer by its index
            slicers.RemoveAt(0);

            // Alternative approach (commented out):
            // Slicer slicerToRemove = slicers[0];
            // slicers.Remove(slicerToRemove);
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}