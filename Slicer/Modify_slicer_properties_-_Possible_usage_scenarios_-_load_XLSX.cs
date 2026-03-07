using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class ModifySlicerProperties
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("Input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve the slicer collection on the worksheet
        SlicerCollection slicers = worksheet.Slicers;

        // Ensure there is at least one slicer to modify
        if (slicers.Count > 0)
        {
            // Get the first slicer (or iterate to find a specific one)
            Slicer slicer = slicers[0];

            // Modify various slicer properties
            slicer.Caption = "Updated Slicer Caption";   // Change the header text
            slicer.ShowCaption = true;                  // Ensure the header is visible
            slicer.NumberOfColumns = 2;                 // Display items in two columns
            slicer.LockedPosition = true;               // Prevent user from moving/resizing the slicer

            // Use the shape's Placement property (the older Placement property is obsolete)
            slicer.Shape.Placement = PlacementType.MoveAndSize;

            // Adjust size in pixels
            slicer.WidthPixel = 250;
            slicer.HeightPixel = 180;

            // Refresh the slicer to apply changes to linked pivot tables
            slicer.Refresh();
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("Output.xlsx", SaveFormat.Xlsx);
    }
}