using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerAutoFormatting
{
    class Program
    {
        static void Main(string[] args)
        {
            // Load the existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                // Access the slicer collection of the current worksheet
                SlicerCollection slicers = worksheet.Slicers;

                // Apply custom formatting to each slicer found
                for (int i = 0; i < slicers.Count; i++)
                {
                    Slicer slicer = slicers[i];

                    // Set a built‑in dark style for the slicer
                    slicer.StyleType = SlicerStyleType.SlicerStyleDark2;

                    // Arrange slicer items in two columns
                    slicer.NumberOfColumns = 2;

                    // Optionally set a custom caption
                    slicer.Caption = "Custom Slicer";
                }
            }

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}