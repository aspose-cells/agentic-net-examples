using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                // Determine workbook path (first argument or default)
                string workbookPath = args.Length > 0 ? args[0] : "sample.xlsx";

                // Prevent FileNotFoundException
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure there is at least one slicer
                if (worksheet.Slicers.Count == 0)
                {
                    Console.WriteLine("No slicers found in the worksheet.");
                    return;
                }

                // Use the first slicer for demonstration
                Slicer slicer = worksheet.Slicers[0];

                // Check if slicer has any selected items
                bool hasSelected = SlicerHelper.HasSelectedItems(slicer);
                Console.WriteLine($"Slicer '{slicer.Name}' has selected items: {hasSelected}");
            }
            catch (Exception ex)
            {
                // Runtime safety
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    public static class SlicerHelper
    {
        // Returns true if the specified slicer has at least one selected item.
        public static bool HasSelectedItems(Slicer slicer)
        {
            // Iterate through all cache items of the slicer.
            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                // The Selected property indicates whether the item is selected.
                if (item.Selected)
                    return true; // Found a selected item.
            }

            // No selected items were found.
            return false;
        }
    }
}