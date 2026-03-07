using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class ModifySlicerProperties
{
    static void Main()
    {
        // Load an existing workbook that contains a slicer
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Access the slicer collection on the worksheet
        SlicerCollection slicers = sheet.Slicers;

        // Ensure there is at least one slicer to modify
        if (slicers.Count > 0)
        {
            // Retrieve the first slicer
            Slicer slicer = slicers[0];

            // Modify various slicer properties
            slicer.Caption = "Modified Slicer Caption";
            slicer.NumberOfColumns = 2;               // Display items in two columns
            slicer.WidthPixel = 250;                  // Set width in pixels
            slicer.HeightPixel = 180;                 // Set height in pixels
            slicer.LockedPosition = true;             // Prevent moving/resizing via UI
            slicer.ShowCaption = false;               // Hide the slicer header
            slicer.StyleType = SlicerStyleType.SlicerStyleLight2; // Apply a built‑in style

            // Refresh the slicer to apply changes to the underlying pivot table
            slicer.Refresh();
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}