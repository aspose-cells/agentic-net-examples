using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace DeleteSlicerSample
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (adjust index or name as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the slicer collection on this worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // If there is at least one slicer, remove the first one (index 0)
            if (slicers.Count > 0)
            {
                slicers.RemoveAt(0); // Deletes the slicer at the specified index
            }

            // Save the workbook after deletion (replace with desired output path)
            workbook.Save("output.xlsx");
        }
    }
}