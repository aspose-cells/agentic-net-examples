using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the slicer collection of the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Set the IsPrintable property of each slicer to false
            foreach (Slicer slicer in slicers)
            {
                slicer.IsPrintable = false;
            }
        }

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}