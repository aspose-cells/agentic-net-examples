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
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Loop backwards to safely remove items while iterating
                for (int i = slicers.Count - 1; i >= 0; i--)
                {
                    Slicer slicer = slicers[i];

                    // Check if the slicer's name starts with "Region"
                    if (!string.IsNullOrEmpty(slicer.Name) && slicer.Name.StartsWith("Region", StringComparison.OrdinalIgnoreCase))
                    {
                        // Remove the slicer at the current index
                        slicers.RemoveAt(i);
                    }
                }
            }

            // Save the modified workbook in XLSX format
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}