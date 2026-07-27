using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main(string[] args)
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the slicer collection of the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Iterate through all slicers and set the printable flag to false
            foreach (Slicer slicer in slicers)
            {
                slicer.IsPrintable = false;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}