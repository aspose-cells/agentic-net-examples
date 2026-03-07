using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class ModifySlicerProperties
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // If there are no slicers, continue to the next worksheet
            if (slicers.Count == 0) continue;

            // Modify each slicer in the collection
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Example property modifications
                slicer.Caption = "Updated Slicer Caption";
                slicer.NumberOfColumns = 2;                     // Show items in two columns
                slicer.LockedPosition = true;                  // Prevent moving/resizing via UI
                slicer.StyleType = SlicerStyleType.SlicerStyleDark1; // Apply a dark style
                slicer.WidthPixel = 250;                       // Set width in pixels
                slicer.HeightPixel = 180;                      // Set height in pixels
                slicer.ShowCaption = true;                     // Ensure the caption is visible
                slicer.ShowAllItems = false;                   // Hide items with no data (optional)

                // Refresh the slicer to apply changes to the underlying pivot table
                slicer.Refresh();
            }
        }

        // Save the modified workbook to a new file (replace with your desired output path)
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}