using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerSelectionDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Values that should be selected in the slicer
                string[] valuesToSelect = { "Apple", "Banana" };

                // Get the first worksheet (adjust if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one slicer
                if (worksheet.Slicers.Count == 0)
                {
                    Console.WriteLine("No slicers found on the first worksheet.");
                    return;
                }

                // Access the first slicer (or modify to select by name)
                Slicer slicer = worksheet.Slicers[0];

                // Update selection for each slicer cache item
                foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
                {
                    item.Selected = valuesToSelect.Contains(item.Value);
                }

                // Apply the changes
                slicer.Refresh();

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}