using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    public class SlicerSelectionHandler
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get slicers on the current worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // Process each slicer
                    for (int s = 0; s < slicers.Count; s++)
                    {
                        Slicer slicer = slicers[s];

                        // Access slicer cache items
                        SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

                        // Log currently selected items
                        Console.WriteLine($"Worksheet: {sheet.Name}, Slicer: {slicer.Name}");
                        for (int i = 0; i < items.Count; i++)
                        {
                            SlicerCacheItem item = items[i];
                            if (item.Selected)
                            {
                                Console.WriteLine($"  Selected Item: {item.Value}");
                            }
                        }

                        // Deselect all items
                        for (int i = 0; i < items.Count; i++)
                        {
                            items[i].Selected = false;
                        }

                        // Refresh slicer to apply changes
                        slicer.Refresh();
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SlicerSelectionHandler.Run();
        }
    }
}