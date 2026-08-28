// Title: Apply a corporate dark slicer style to all slicers across every worksheet in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file, iterates through each worksheet and each slicer, sets the slicer style to SlicerStyleDark1, configures 2 columns, width 250 px, height 180 px, locks the position, enables caption and Show All Items, then saves the workbook. | Write a reusable method that accepts a Workbook object and slicer formatting parameters (style, column count, width, height, locked flag, caption visibility, show‑all‑items flag) and applies them to every slicer in every worksheet via the Aspose.Cells API.
// Common Searches: how to programmatically set slicer style and dimensions for all worksheets using Aspose.Cells C# | Aspose.Cells bulk update slicer properties like width, height, and locked position in an Excel workbook | C# code to apply a corporate dark slicer theme to every slicer in an Excel file with Aspose.Cells | iterate through worksheets and modify slicer settings with Aspose.Cells for .NET | set slicer ShowAllItems and ShowCaption for all slicers using Aspose.Cells API
// Tags: Aspose.Cells bulk slicer style update | C# iterate worksheets to modify slicers | set slicer width and height Aspose.Cells | lock slicer position Excel Aspose.Cells | apply corporate dark slicer theme .NET

using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Drawing;

// The example loads input.xlsx, defines corporate slicer settings (dark style, 2 columns, 250 px width, 180 px height, locked position, caption visible, Show All Items enabled), iterates through every worksheet and each slicer to apply those settings—including locking the underlying shape—and saves the modified workbook as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the workbook that contains slicers
        Workbook workbook = new Workbook("input.xlsx");

        // Corporate style settings
        SlicerStyleType corporateStyle = SlicerStyleType.SlicerStyleDark1;
        int corporateColumns = 2;
        int corporateWidthPixel = 250;
        int corporateHeightPixel = 180;
        bool corporateLockedPosition = true;
        bool corporateShowCaption = true;
        bool corporateShowAllItems = true;

        // Apply the style to every slicer in every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            SlicerCollection slicers = sheet.Slicers;

            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                slicer.StyleType = corporateStyle;
                slicer.NumberOfColumns = corporateColumns;
                slicer.WidthPixel = corporateWidthPixel;
                slicer.HeightPixel = corporateHeightPixel;
                slicer.LockedPosition = corporateLockedPosition;
                slicer.ShowCaption = corporateShowCaption;
                slicer.ShowAllItems = corporateShowAllItems;

                // Ensure the underlying shape is also locked
                slicer.Shape.IsLocked = corporateLockedPosition;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
