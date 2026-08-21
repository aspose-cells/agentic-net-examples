// Title: C# – Reset All Slicer Selections in an Excel Workbook with Aspose.Cells
// Description: This example demonstrates how to load an existing or new workbook, iterate through every slicer on the first worksheet, deselect all items in each slicer's cache, refresh the slicer, and save the updated file. It provides a reliable way to return slicer filters to their default (unselected) state programmatically.
// Keywords: Aspose.Cells | C# slicer reset | clear slicer selections | Excel slicer programmatic | reset slicer filter .NET | SlicerCacheItem Selected false | refresh slicer Aspose | Excel automation | workbook cleanup
// Common Searches: how to clear slicer selections with Aspose.Cells | reset Excel slicer filter using C# | programmatically deselect all slicer items .NET | Aspose.Cells example to refresh slicer after clearing | remove slicer selections before saving workbook
// Developer Intent: Programmatically clear every selected item in all slicers of a worksheet to restore the default filter state.
// Use Cases: Prepare a report workbook so all slicers start unfiltered before distribution. | Automate cleanup after data refresh by resetting slicer filters. | Ensure a shared template opens with no slicer selections applied.
// AI Prompts: Generate C# code using Aspose.Cells that clears all slicer selections, refreshes each slicer, and saves the workbook. | Explain how to safely reset slicer filters in an existing Excel file when no slicers are present. | Show a minimal Aspose.Cells snippet to deselect slicer items without altering other worksheet data.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for slicer related classes

namespace AsposeCellsExamples
{
    // This example demonstrates how to load an existing or new workbook, iterate through every slicer on the first worksheet, deselect all items in each slicer's cache, refresh the slicer, and save the updated file. It provides a reliable way to return slicer filters to their default (unselected) state programmatically.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "Output.xlsx";

                Workbook workbook;

                // Load existing workbook if the file exists; otherwise create a new one
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook();
                }

                // Get the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Access the slicer collection on the worksheet
                SlicerCollection slicers = worksheet.Slicers;

                // Iterate through each slicer and clear its selections
                foreach (Slicer slicer in slicers)
                {
                    // Get all items in the slicer's cache
                    SlicerCacheItemCollection cacheItems = slicer.SlicerCache.SlicerCacheItems;

                    // Deselect every item
                    for (int i = 0; i < cacheItems.Count; i++)
                    {
                        cacheItems[i].Selected = false;
                    }

                    // Refresh the slicer to apply the changes
                    slicer.Refresh();
                }

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
