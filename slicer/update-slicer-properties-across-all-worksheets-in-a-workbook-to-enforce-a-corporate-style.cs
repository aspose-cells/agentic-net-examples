using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

class UpdateSlicerStyle
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("InputWorkbook.xlsx");

        // Define corporate style settings
        SlicerStyleType corporateStyle = SlicerStyleType.SlicerStyleDark2;
        int corporateColumns = 2;
        int corporateWidthPixel = 250;
        int corporateHeightPixel = 150;
        bool corporateLockedPosition = true;
        bool corporateShowCaption = true;
        bool corporateShowAllItems = false;

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the slicer collection of the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Apply corporate style to each slicer on the worksheet
            for (int i = 0; i < slicers.Count; i++)
            {
                Slicer slicer = slicers[i];

                // Set the desired properties
                slicer.StyleType = corporateStyle;
                slicer.NumberOfColumns = corporateColumns;
                slicer.WidthPixel = corporateWidthPixel;
                slicer.HeightPixel = corporateHeightPixel;
                slicer.LockedPosition = corporateLockedPosition;
                slicer.ShowCaption = corporateShowCaption;
                slicer.ShowAllItems = corporateShowAllItems;

                // Refresh the slicer to apply changes to linked PivotTables
                slicer.Refresh();
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("OutputWorkbook.xlsx");
    }
}