using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class DeleteSlicer
{
    static void Main()
    {
        // Load the existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the slicer (e.g., the first worksheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = sheet.Slicers;

        // If there is at least one slicer, remove the first one (index 0)
        if (slicers.Count > 0)
        {
            slicers.RemoveAt(0);
        }

        // Save the workbook after removing the slicer
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}