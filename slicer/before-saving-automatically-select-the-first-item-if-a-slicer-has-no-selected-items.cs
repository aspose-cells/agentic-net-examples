// Title: How to automatically select the first slicer item when none are selected before saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that loops through every worksheet's slicers, checks each slicer's cache, and sets IsSelected = true on the first slicer item if the slicer has no selected items before calling Workbook.Save. | Update an existing Aspose.Cells .NET application to ensure each slicer automatically selects its first item when the slicer cache reports no selections, with per‑slicer error handling.
// Common Searches: Aspose.Cells C# select default slicer item when slicer is empty | Set first slicer item as selected in Excel workbook using Aspose.Cells before saving | C# loop through workbook slicers and ensure a selection exists with Aspose.Cells | How to handle slicer cache with no selected items in Aspose.Cells .NET
// Tags: Aspose.Cells slicer cache default selection | Aspose.Cells workbook slicer iteration | set slicer item IsSelected programmatically | auto-select first slicer item before workbook save | per-slicer error handling Aspose.Cells

using System;
using System.IO;
using System.Linq;
using System.Collections;
using Aspose.Cells;
using Aspose.Cells.Slicers; // Required for Slicer, SlicerCache, and SlicerItem classes

namespace AsposeCellsSlicerExample
{
    // The example loads an Excel file, iterates over each worksheet and its slicers, checks whether any slicer items are selected, and if none are, marks the first slicer item as selected before saving the workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all slicers on the worksheet
                    foreach (Slicer slicer in worksheet.Slicers)
                    {
                        try
                        {
                            // Use dynamic to access slicer cache items (avoids compile‑time binding issues)
                            dynamic cacheDyn = slicer.SlicerCache;
                            if (cacheDyn == null)
                                continue; // No cache, skip this slicer

                            IEnumerable items = cacheDyn.SlicerItems as IEnumerable;
                            if (items == null)
                                continue; // No items to process

                            // Determine if any item is selected
                            bool anyItemSelected = false;
                            foreach (dynamic item in items)
                            {
                                if (item.IsSelected)
                                {
                                    anyItemSelected = true;
                                    break;
                                }
                            }

                            // If no items are selected, select the first item
                            if (!anyItemSelected)
                            {
                                var enumerator = items.GetEnumerator();
                                if (enumerator.MoveNext())
                                {
                                    dynamic firstItem = enumerator.Current;
                                    firstItem.IsSelected = true;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            // Handle errors related to a specific slicer without stopping the whole process
                            Console.WriteLine($"Error processing slicer '{slicer.Name}': {ex.Message}");
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
