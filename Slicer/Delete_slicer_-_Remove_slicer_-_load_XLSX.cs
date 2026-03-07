using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDeletion
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook that contains slicers
            Workbook workbook = new Workbook("input.xlsx");

            // Assume the slicer to delete is on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the slicer collection of the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Check if there is at least one slicer
            if (slicers.Count > 0)
            {
                // Remove the first slicer (index 0) from the collection
                slicers.RemoveAt(0);
            }

            // Save the modified workbook
            workbook.Save("output.xlsx");
        }
    }
}