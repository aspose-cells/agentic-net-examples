using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class SlicerSelectionDemo
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sheet.Slicers == null || sheet.Slicers.Count == 0)
                throw new InvalidOperationException("No slicers found on the first worksheet.");

            // Retrieve the first slicer
            Slicer slicer = sheet.Slicers[0];

            // Select desired items ("Apple" and "Banana") in the slicer
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                item.Selected = item.Value == "Apple" || item.Value == "Banana";
            }

            // Refresh the slicer so the connected pivot table reflects the new selections
            slicer.Refresh();

            // Save the updated workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}