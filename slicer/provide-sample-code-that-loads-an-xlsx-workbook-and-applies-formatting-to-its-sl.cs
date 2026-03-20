using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class SlicerFormattingDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook that already contains slicers
        Workbook workbook = new Workbook("InputWithSlicers.xlsx");

        // Access the worksheet where slicers are placed (here we use the first sheet)
        Worksheet sheet = workbook.Worksheets[0];

        // Loop through each slicer in the worksheet and apply formatting
        foreach (Slicer slicer in sheet.Slicers)
        {
            // Apply a built‑in dark style
            slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

            // Set a custom caption that includes the slicer name
            slicer.Caption = $"Filter: {slicer.Name}";
            slicer.ShowCaption = true; // Ensure the caption header is visible

            // Arrange slicer items in two columns
            slicer.NumberOfColumns = 2;

            // Lock the slicer position so users cannot move or resize it
            slicer.LockedPosition = true;

            // Adjust the size of the slicer (pixel units)
            slicer.WidthPixel = 200;
            slicer.HeightPixel = 120;
        }

        // Save the workbook with the updated slicer formatting
        workbook.Save("OutputFormattedSlicers.xlsx");
    }
}