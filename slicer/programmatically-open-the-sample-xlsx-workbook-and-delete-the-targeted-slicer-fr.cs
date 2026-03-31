using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the existing workbook (replace with the actual path to your sample file)
            Workbook workbook = new Workbook("Sample.xlsx");

            // Assume the slicer is on the first worksheet; adjust the index if needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the slicer collection for the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Define the name of the slicer you want to delete
            string slicerNameToDelete = "FruitSlicer";

            // Locate the slicer by name
            Slicer targetSlicer = null;
            foreach (Slicer slicer in slicers)
            {
                if (slicer.Name.Equals(slicerNameToDelete, StringComparison.OrdinalIgnoreCase))
                {
                    targetSlicer = slicer;
                    break;
                }
            }

            // If the slicer was found, remove it from the collection
            if (targetSlicer != null)
            {
                slicers.Remove(targetSlicer);
                Console.WriteLine($"Slicer \"{slicerNameToDelete}\" has been removed.");
            }
            else
            {
                Console.WriteLine($"Slicer \"{slicerNameToDelete}\" was not found.");
            }

            // Save the modified workbook
            workbook.Save("Sample_modified.xlsx");
        }
    }
}