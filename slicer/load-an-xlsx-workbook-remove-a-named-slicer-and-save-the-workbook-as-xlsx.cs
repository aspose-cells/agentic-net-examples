using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Specify the name of the slicer to remove
            string slicerNameToRemove = "MySlicer";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Find the slicer with the specified name
                for (int i = 0; i < slicers.Count; i++)
                {
                    Slicer slicer = slicers[i];
                    if (slicer.Name == slicerNameToRemove)
                    {
                        // Remove the slicer from the collection
                        slicers.Remove(slicer);
                        // Exit loops after removal
                        break;
                    }
                }
            }

            // Save the modified workbook as XLSX
            workbook.Save("output.xlsx");
        }
    }
}