// Title: How to deselect slicer cache items containing a specific keyword in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that opens a workbook, finds all slicers, and sets Selected = false for any slicer cache item whose text includes a given keyword (case‑insensitive). | Create a reusable method that accepts an input file path, output file path, and keyword, then iterates through each worksheet's slicer caches to deselect matching items and saves the modified workbook. | Show how to add robust error handling for missing files, absent slicers, or empty slicer caches while performing keyword‑based deselection of slicer items in a .NET application.
// Common Searches: Aspose.Cells C# deselect slicer items that contain a keyword | Programmatically unselect slicer cache entries in Excel using .NET | Iterate through slicer cache items and set Selected false based on text value Aspose.Cells | How to filter slicer selections by keyword in a .xlsx file with Aspose.Cells | Remove specific slicer selections from an Excel workbook via C#
// Tags: slicer cache item deselection Aspose.Cells | keyword based slicer filtering .NET | iterate slicer cache items C# | update slicer selections Excel workbook | save modified workbook Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.IO;

// // Loads an Excel workbook, walks through every worksheet and its slicers, deselects any slicer cache items whose value contains the specified keyword (case‑insensitive), and saves the updated workbook to a new file.
class SlicerCacheDeselector
{
    static void Main()
    {
        // Define file paths
        string inputPath = @"C:\Input\Sample.xlsx";
        string outputPath = @"C:\Output\Sample_Modified.xlsx";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Keyword to match slicer items that should be deselected
            string keyword = "Exclude";

            // Iterate through all worksheets and their slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                var slicers = sheet.Slicers;
                if (slicers == null) continue;

                foreach (Slicer slicer in slicers)
                {
                    var cache = slicer.SlicerCache;
                    if (cache == null) continue;

                    var items = cache.SlicerCacheItems;
                    if (items == null) continue;

                    // Iterate over each slicer cache item
                    foreach (SlicerCacheItem item in items)
                    {
                        // Use the item's value (as string) for comparison
                        string itemValue = item.Value?.ToString();
                        if (!string.IsNullOrEmpty(itemValue) &&
                            itemValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Deselect matching items
                            item.Selected = false;
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outDir) && !Directory.Exists(outDir))
                Directory.CreateDirectory(outDir);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
