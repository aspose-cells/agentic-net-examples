using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the slicer (e.g., the first worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Specify the name of the slicer you want to remove
        string slicerNameToRemove = "MySlicer";

        // Locate the slicer object by its name
        Slicer slicerToRemove = null;
        foreach (Slicer slicer in slicers)
        {
            if (slicer.Name.Equals(slicerNameToRemove, StringComparison.OrdinalIgnoreCase))
            {
                slicerToRemove = slicer;
                break;
            }
        }

        // If the slicer was found, remove it from the collection
        if (slicerToRemove != null)
        {
            slicers.Remove(slicerToRemove);
        }
        else
        {
            Console.WriteLine($"Slicer \"{slicerNameToRemove}\" not found.");
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}