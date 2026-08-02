using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Iterate through each slicer and log its name
        foreach (Slicer slicer in slicers)
        {
            Console.WriteLine("Slicer Name: " + slicer.Name);
        }

        // Save the workbook if any changes were made (optional)
        workbook.Save("output.xlsx");
    }
}