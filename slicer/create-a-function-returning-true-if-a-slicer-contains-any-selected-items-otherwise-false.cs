using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerUtilities
{
    public static class SlicerHelper
    {
        /// <summary>
        /// Returns true if the specified slicer has at least one selected item.
        /// </summary>
        public static bool HasSelectedItems(Slicer slicer)
        {
            if (slicer == null) throw new ArgumentNullException(nameof(slicer));

            foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
            {
                if (item.Selected)
                    return true;
            }

            return false;
        }

        /// <summary>
        /// Demonstrates loading a workbook and checking slicer selections.
        /// </summary>
        public static void Demo()
        {
            const string filePath = "InputWithSlicer.xlsx";

            // Prevent FileNotFoundException
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load workbook
                Workbook workbook = new Workbook(filePath);

                // Ensure a slicer exists on the first worksheet
                if (workbook.Worksheets[0].Slicers.Count == 0)
                {
                    Console.WriteLine("No slicers found on the first worksheet.");
                    return;
                }

                // Get the first slicer
                Slicer slicer = workbook.Worksheets[0].Slicers[0];

                // Check for selected items
                bool anySelected = HasSelectedItems(slicer);
                Console.WriteLine($"Slicer '{slicer.Name}' has selected items: {anySelected}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                SlicerHelper.Demo();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}