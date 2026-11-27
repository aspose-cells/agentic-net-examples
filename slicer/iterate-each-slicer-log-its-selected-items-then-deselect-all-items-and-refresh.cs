// Title: How to enumerate slicer items, log selected values, clear selections, and refresh slicers in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# iterate over slicer cache items and get selected values | C# code to clear all selections in Excel slicers with Aspose.Cells | Refresh slicer after changing selection state using Aspose.Cells .NET | Log selected slicer items from each worksheet in a workbook with Aspose.Cells | How to programmatically deselect slicer items and save workbook using Aspose.Cells
// Tags: enumerate slicer cache items Aspose.Cells | output selected slicer items C# | clear slicer selections Aspose.Cells | refresh slicer programmatically .NET | handle slicers across multiple worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an Excel workbook, iterates through every worksheet and its slicers, logs each item whose Selected property is true, clears the selection for all items, refreshes the slicer to apply the changes, and saves the updated workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

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

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all slicers on the worksheet
                foreach (Slicer slicer in sheet.Slicers)
                {
                    // Access the slicer's cache
                    SlicerCache cache = slicer.SlicerCache;

                    // Log each selected item
                    foreach (SlicerCacheItem item in cache.SlicerCacheItems)
                    {
                        // In Aspose.Cells, selection state is accessed via Selected property
                        if (item.Selected)
                        {
                            // The displayed text of the item is available via the Value property
                            Console.WriteLine($"Worksheet: {sheet.Name}, Slicer: {slicer.Name}, Selected Item: {item.Value}");
                        }
                    }

                    // Deselect all items in the slicer
                    foreach (SlicerCacheItem item in cache.SlicerCacheItems)
                    {
                        item.Selected = false;
                    }

                    // Refresh the slicer to apply changes
                    slicer.Refresh();
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
