// Title: Programmatically deselect slicer items containing a specific keyword in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# deselect slicer items that contain a keyword | iterate through SlicerCacheItems and unselect matching values in Excel with Aspose.Cells | remove slicer selections based on text in a workbook using Aspose.Cells .NET | C# code to filter slicer cache items by keyword and set Selected to false | how to clear specific slicer options in an Excel file with Aspose.Cells
// Tags: slicer cache deselection Aspose.Cells | keyword based slicer filtering C# | Excel slicer automation .NET | programmatic slicer selection control | iterate slicer cache items Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an existing XLSX workbook with Aspose.Cells, walks through each worksheet's slicers, examines every SlicerCacheItem, and when an item's value contains the given keyword (case‑insensitive) it sets the item's Selected property to false. The modified workbook is then saved to a new file.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";
        string keyword = "YourKeyword";

        try
        {
            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and their slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    SlicerCache cache = slicer.SlicerCache;
                    if (cache == null) continue;

                    // Iterate over each slicer cache item
                    foreach (SlicerCacheItem item in cache.SlicerCacheItems)
                    {
                        // Check if the item's value contains the keyword (case‑insensitive)
                        if (item.Value != null &&
                            item.Value.ToString().IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            // Deselect the matching item
                            item.Selected = false;
                        }
                    }
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
