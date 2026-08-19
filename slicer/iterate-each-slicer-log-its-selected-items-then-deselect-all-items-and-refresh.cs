// Title: C# – List Selected Slicer Items, Deselect All, and Refresh Workbook with Aspose.Cells
// Description: Loads an Excel file, iterates every worksheet and slicer, logs each selected cache item, clears all selections, refreshes the slicer, and saves the updated workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells slicer C# | list selected slicer items | deselect slicer programmatically | refresh slicer Aspose | iterate workbook slicers | Excel slicer automation .NET | global Excel slicer handling
// Common Searches: how to read slicer selections with Aspose.Cells | C# code to clear all slicer items in Excel | refresh slicer after changing selections Aspose | iterate through slicers in a workbook using .NET | log slicer filter values before resetting
// Developer Intent: Programmatically enumerate each slicer, output its selected values, reset all selections, and apply the changes.
// Use Cases: Create a pre‑export audit of slicer filters before distributing a workbook. | Automate workbook sanitization by clearing slicer selections across all sheets. | Capture slicer state for debugging or logging prior to saving modifications.
// AI Prompts: Generate C# code with Aspose.Cells that prints all selected slicer items, deselects them, refreshes each slicer, and saves the workbook. | Show how to handle slicers that have no selected items when logging and clearing selections in a .NET Excel automation script. | Explain step‑by‑step how to iterate over a workbook’s slicer collections, log selected values, clear selections, and refresh the slicer using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // Loads an Excel file, iterates every worksheet and slicer, logs each selected cache item, clears all selections, refreshes the slicer, and saves the updated workbook using Aspose.Cells for .NET.
    public class SlicerOperations
    {
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get slicer collection for the current worksheet
                    SlicerCollection slicers = sheet.Slicers;

                    // Process each slicer
                    for (int s = 0; s < slicers.Count; s++)
                    {
                        Slicer slicer = slicers[s];

                        // Access slicer cache items
                        SlicerCacheItemCollection items = slicer.SlicerCache.SlicerCacheItems;

                        // Log selected items
                        Console.WriteLine($"Worksheet: {sheet.Name}, Slicer: {slicer.Name}");
                        for (int i = 0; i < items.Count; i++)
                        {
                            SlicerCacheItem item = items[i];
                            if (item.Selected)
                            {
                                Console.WriteLine($"  Selected Item: {item.Value}");
                            }

                            // Deselect the item
                            item.Selected = false;
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
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
