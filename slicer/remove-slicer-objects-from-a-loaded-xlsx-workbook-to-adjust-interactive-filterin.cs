using Aspose.Cells;
using Aspose.Cells.Slicers;

class RemoveSlicersDemo
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all worksheets in the workbook
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = ws.Slicers;

            // Remove all slicers from this worksheet
            slicers.Clear();
        }

        // Save the workbook after slicers have been removed
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}