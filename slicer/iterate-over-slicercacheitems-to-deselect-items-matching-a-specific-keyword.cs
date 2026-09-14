// Title: Deselect Excel slicer items that match a given keyword using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, refreshes slicer caches, and sets Selected = false for any slicer cache item whose Value matches a given keyword. | Create a method that walks through each worksheet and slicer, clears the selection of entries whose label is "Apple", refreshes the slicer, and saves the updated file. | Generate a script that programmatically changes slicer selections based on a keyword and writes the result to a new Excel workbook using Aspose.Cells.
// Common Searches: how to filter slicer entries by text using Aspose.Cells C# | C# code to unselect specific slicer values in Excel | programmatically clear slicer selection for a given keyword with Aspose.Cells | iterate through slicer cache items and change selection in .NET | Aspose.Cells example to modify slicer selections based on a string
// Tags: Aspose.Cells deselect slicer items | C# loop over slicer cache items | slicer cache refresh Aspose.Cells | Excel slicer selection programmatic | unselect slicer entries by keyword

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsExamples
{
    // // Loads an Excel workbook, refreshes all slicer caches, iterates through each worksheet and slicer, deselects any slicer cache item whose value equals the keyword "Apple", refreshes the slicer, and saves the modified workbook.
    public class DeselectSlicerItemsByKeyword
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook that contains the slicer
            Workbook workbook = new Workbook(inputPath);

            // Refresh all slicer caches
            workbook.RefreshAll();

            // Keyword of items to be deselected
            string keyword = "Apple";

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through each slicer on the worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Get the collection of slicer cache items
                    SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

                    // Deselect items whose label matches the keyword
                    foreach (SlicerCacheItem item in items)
                    {
                        if (item.Value == keyword)
                        {
                            item.Selected = false;
                        }
                    }

                    // Refresh the slicer to apply the changes
                    slicer.Refresh();
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
    }
}
