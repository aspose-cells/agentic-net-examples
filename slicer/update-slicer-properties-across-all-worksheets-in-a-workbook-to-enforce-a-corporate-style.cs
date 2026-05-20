using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

class UpdateSlicerStyle
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Define corporate style settings
        SlicerStyleType corporateStyle = SlicerStyleType.SlicerStyleDark2;
        int corporateColumns = 2;
        int corporateWidthPixel = 250;
        int corporateHeightPixel = 150;
        bool corporateLockedPosition = true;
        bool corporateShowCaption = true;

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the slicer collection of the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Apply corporate style to each slicer
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Set built‑in style
                slicer.StyleType = corporateStyle;

                // Set number of columns
                slicer.NumberOfColumns = corporateColumns;

                // Set size in pixels
                slicer.WidthPixel = corporateWidthPixel;
                slicer.HeightPixel = corporateHeightPixel;

                // Lock position to prevent user moving/resizing
                slicer.LockedPosition = corporateLockedPosition;

                // Ensure caption visibility matches corporate policy
                slicer.ShowCaption = corporateShowCaption;

                // Refresh slicer to apply changes to linked pivot tables
                slicer.Refresh();
            }
        }

        // Save the modified workbook
        workbook.Save("OutputWorkbook.xlsx");
    }
}