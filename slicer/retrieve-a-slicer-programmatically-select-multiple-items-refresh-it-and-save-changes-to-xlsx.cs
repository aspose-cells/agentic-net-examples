// Title: Programmatically deselect all items, select specific values in an Excel slicer, recalculate formulas, and save the workbook using Aspose.Cells for .NET
// AI Prompts: Load an XLSX file with Aspose.Cells, locate the first slicer on the first worksheet, clear its selections, mark the items "Item1", "Item2", and "Item3" as selected, recalculate the workbook, and write the result to a new file. | Using Aspose.Cells SlicerCache, iterate over SlicerCacheItem objects to set Selected = false for every item, then set Selected = true for a predefined list of values, and persist the changes to a new XLSX file. | Refresh an Excel slicer after modifying its cache items by invoking Workbook.CalculateFormula and then saving the workbook in XLSX format with Aspose.Cells.
// Common Searches: How to change slicer selections in an Excel file using Aspose.Cells C# | Aspose.Cells programmatically select multiple slicer items and update workbook | C# code to deselect all slicer items then select specific ones with Aspose.Cells | Refresh slicer after modifying cache items in Aspose.Cells .NET | Save Excel workbook after slicer filter changes using Aspose.Cells
// Tags: Aspose.Cells slicer cache item selection | C# deselect all slicer items Aspose | select multiple slicer values .NET | recalculate workbook after slicer update Aspose | save XLSX after slicer modification Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.Collections.Generic;
using System.IO;

// // Loads an existing XLSX, retrieves the first slicer, clears all selections, selects "Item1", "Item2", "Item3" via the slicer cache, recalculates formulas, and saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one slicer
            if (worksheet.Slicers.Count == 0)
            {
                Console.WriteLine("No slicers found on the worksheet.");
                workbook.Save(outputPath, SaveFormat.Xlsx);
                return;
            }

            // Retrieve the first slicer
            Slicer slicer = worksheet.Slicers[0];

            // Get the slicer cache that holds the items
            SlicerCache slicerCache = slicer.SlicerCache;

            // Deselect all items to start from a clean state
            foreach (SlicerCacheItem cacheItem in slicerCache.SlicerCacheItems)
            {
                cacheItem.Selected = false;
            }

            // Define the items you want to select
            List<string> itemsToSelect = new List<string> { "Item1", "Item2", "Item3" };

            // Select the specified items
            foreach (SlicerCacheItem cacheItem in slicerCache.SlicerCacheItems)
            {
                if (itemsToSelect.Contains(cacheItem.Value))
                {
                    cacheItem.Selected = true;
                }
            }

            // Recalculate formulas so the slicer reflects the new selections
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
