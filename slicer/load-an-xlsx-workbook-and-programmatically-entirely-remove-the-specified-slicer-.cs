using Aspose.Cells;
using Aspose.Cells.Slicers;

class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the worksheet that contains the slicer (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Specify the name of the slicer you want to remove
        string slicerNameToRemove = "MySlicer";

        // Get the slicer collection from the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Locate the slicer by its Name property
        Slicer slicerToDelete = null;
        foreach (Slicer slicer in slicers)
        {
            if (slicer.Name == slicerNameToRemove)
            {
                slicerToDelete = slicer;
                break;
            }
        }

        // If the slicer was found, remove it from the collection
        if (slicerToDelete != null)
        {
            slicers.Remove(slicerToDelete);
        }

        // Save the workbook after removal
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}