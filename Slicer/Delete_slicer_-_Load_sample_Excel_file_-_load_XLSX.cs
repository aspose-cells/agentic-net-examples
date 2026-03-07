using Aspose.Cells;
using Aspose.Cells.Slicers;

class DeleteSlicerExample
{
    static void Main()
    {
        // Load the existing XLSX file
        Workbook workbook = new Workbook("Sample.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Retrieve the slicer collection on this worksheet
        SlicerCollection slicers = sheet.Slicers;

        // Remove the first slicer if any exist
        if (slicers.Count > 0)
        {
            slicers.RemoveAt(0);
        }

        // Save the workbook after the slicer has been removed
        workbook.Save("Sample_NoSlicer.xlsx", SaveFormat.Xlsx);
    }
}