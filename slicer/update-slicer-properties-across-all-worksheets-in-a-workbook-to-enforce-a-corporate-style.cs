// Title: Apply a Corporate Slicer Style to Every Worksheet with Aspose.Cells for .NET
// Description: Loads an Excel workbook, defines a corporate slicer theme (style, column count, width, height, locked position, caption), iterates through all worksheets, updates each slicer with those settings, refreshes them, and saves the modified file.
// Keywords: Aspose.Cells slicer style .NET | bulk slicer formatting | set slicer width height C# | lock slicer position programmatically | apply corporate slicer theme | update slicer caption Aspose | global Excel dashboard styling
// Common Searches: How to apply the same slicer style to all worksheets using Aspose.Cells | Programmatically change slicer width, height and columns in a .NET Excel file | Lock slicer position and set a uniform caption with Aspose.Cells | Bulk update slicer properties across multiple sheets in C# | Standardize slicer appearance for corporate reports in Excel
// Developer Intent: Modify every slicer in a workbook so it follows a predefined corporate style and layout.
// Use Cases: Ensure consistent slicer look‑and‑feel across multi‑sheet financial dashboards before distribution. | Prevent end‑users from moving or resizing slicers in shared reports by locking their position. | Add a uniform corporate caption to slicers generated from pivot tables on several worksheets.
// AI Prompts: Write C# code with Aspose.Cells that loops through all worksheets and sets slicer style, column count, width, height, locked position, and caption for each slicer, then saves the workbook. | Create a reusable method that accepts slicer style parameters and applies them to every slicer in a given workbook using Aspose.Cells. | Explain how to read existing slicer properties first, then update only those that differ from the corporate defaults.

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

// Loads an Excel workbook, defines a corporate slicer theme (style, column count, width, height, locked position, caption), iterates through all worksheets, updates each slicer with those settings, refreshes them, and saves the modified file.
class UpdateSlicerStyle
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Define corporate style settings
        SlicerStyleType corporateStyle = SlicerStyleType.SlicerStyleDark2; // example corporate style
        int corporateColumns = 3;
        int corporateWidthPixel = 250;
        int corporateHeightPixel = 150;
        bool corporateLockedPosition = true;
        string corporateCaption = "Corporate Slicer";

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the slicer collection for the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Apply corporate style to each slicer
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Set style and layout properties
                slicer.StyleType = corporateStyle;
                slicer.NumberOfColumns = corporateColumns;
                slicer.WidthPixel = corporateWidthPixel;
                slicer.HeightPixel = corporateHeightPixel;

                // Lock slicer position to prevent user moving/resizing
                slicer.LockedPosition = corporateLockedPosition;

                // Set a uniform caption (optional)
                slicer.Caption = corporateCaption;

                // Refresh the slicer to apply changes
                slicer.Refresh();
            }
        }

        // Save the workbook with updated slicer properties
        workbook.Save("OutputWorkbook.xlsx");
    }
}
