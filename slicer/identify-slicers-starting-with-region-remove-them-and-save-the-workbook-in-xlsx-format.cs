using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerRemoval
{
    class Program
    {
        static void Main()
        {
            // Load the workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Loop backwards so that removal does not affect the index order
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    // Check if the slicer's name starts with "Region"
                    if (slicers[i].Name.StartsWith("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the slicer at the current index
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook in XLSX format
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}