using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class SlicerSelectionExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume the slicer is on the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (sheet.Slicers == null || sheet.Slicers.Count == 0)
                throw new InvalidOperationException("No slicers found on the first worksheet.");

            // Retrieve the first slicer (adjust index or name as needed)
            Slicer slicer = sheet.Slicers[0];

            // Access the slicer cache items collection
            SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

            // Values to be selected in the slicer
            string[] valuesToSelect = { "Apple", "Banana" };

            // Update selection state for each slicer item
            for (int i = 0; i < items.Count; i++)
            {
                SlicerCacheItem item = items[i];
                bool shouldSelect = Array.Exists(
                    valuesToSelect,
                    v => v.Equals(item.Value?.ToString(), StringComparison.OrdinalIgnoreCase));

                item.Selected = shouldSelect;
            }

            // Apply the selection and refresh the connected pivot table
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