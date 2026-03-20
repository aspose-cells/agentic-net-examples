using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerSettingsUpdate
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Assume slicers are placed on the first worksheet; adjust index as needed
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the slicer collection from the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // Iterate through all slicers and modify their properties
            foreach (Slicer slicer in slicers)
            {
                // Update the caption displayed on the slicer
                slicer.Caption = "Updated Slicer Caption";

                // Change the number of columns displayed in the slicer
                slicer.NumberOfColumns = 2;

                // Allow the slicer to be moved/resized by the user
                slicer.LockedPosition = false;

                // Apply a built‑in slicer style (requires Aspose.Cells.Slicers namespace)
                slicer.StyleType = SlicerStyleType.SlicerStyleLight2;

                // Refresh the slicer to ensure any linked PivotTables are recalculated
                slicer.Refresh();
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}