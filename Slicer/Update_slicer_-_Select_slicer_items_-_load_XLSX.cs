using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Slicers;

namespace UpdateSlicerSelection
{
    class Program
    {
        static void Main()
        {
            // Define the input workbook path (adjust as needed)
            string inputPath = Path.Combine(Environment.CurrentDirectory, "InputWithSlicer.xlsx");
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook that contains a slicer
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet; adjust as needed
            Worksheet sheet = workbook.Worksheets[0];

            // Get the slicer collection from the worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Ensure there is at least one slicer
            if (slicers.Count == 0)
            {
                Console.WriteLine("No slicers found in the worksheet.");
                return;
            }

            // Retrieve the first slicer (or use slicers["SlicerName"] if you know the name)
            Slicer slicer = slicers[0];

            // Access the slicer cache items
            SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

            // Example: select the first two items and deselect the rest
            for (int i = 0; i < items.Count; i++)
            {
                SlicerCacheItem item = items[i];
                item.Selected = i < 2;
            }

            // Refresh the slicer so that the underlying PivotTable reflects the new selection
            slicer.Refresh();

            // Save the workbook with the updated slicer selection
            string outputPath = Path.Combine(Environment.CurrentDirectory, "OutputWithUpdatedSlicer.xlsx");
            workbook.Save(outputPath);

            Console.WriteLine($"Slicer selection updated and workbook saved to: {outputPath}");
        }
    }
}