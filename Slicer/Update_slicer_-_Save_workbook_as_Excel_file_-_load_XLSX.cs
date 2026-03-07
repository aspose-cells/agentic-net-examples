using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerUpdateExample
{
    class Program
    {
        static void Main()
        {
            // Paths for the source and output Excel files
            string sourcePath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Load the existing workbook from the source file
            Workbook workbook = new Workbook(sourcePath);

            // Assume the slicer is placed on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // If there is at least one slicer, refresh it (updates the slicer and its PivotTable)
            if (sheet.Slicers.Count > 0)
            {
                Slicer slicer = sheet.Slicers[0];
                slicer.Refresh();
            }

            // Save the modified workbook as an XLSX file
            workbook.Save(outputPath, SaveFormat.Xlsx);

            // Load the newly saved workbook to verify the changes
            Workbook reloadedWorkbook = new Workbook(outputPath);

            // Simple verification: output the number of worksheets
            Console.WriteLine($"Reloaded workbook contains {reloadedWorkbook.Worksheets.Count} worksheet(s).");
        }
    }
}