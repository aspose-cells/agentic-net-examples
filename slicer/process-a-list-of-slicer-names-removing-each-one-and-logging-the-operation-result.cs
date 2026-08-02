using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.Collections.Generic;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Assume slicers are on the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        SlicerCollection slicers = worksheet.Slicers;

        // List of slicer names to be removed
        List<string> slicerNames = new List<string> { "Slicer1", "Slicer2", "Slicer3" };

        foreach (string name in slicerNames)
        {
            Slicer slicer = null;
            try
            {
                // Attempt to get the slicer by name
                slicer = slicers[name];
            }
            catch (Exception)
            {
                // slicer not found, slicer remains null
            }

            if (slicer != null)
            {
                // Remove the slicer from the collection
                slicers.Remove(slicer);
                Console.WriteLine($"Removed slicer '{name}'.");
            }
            else
            {
                Console.WriteLine($"Slicer '{name}' not found.");
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}