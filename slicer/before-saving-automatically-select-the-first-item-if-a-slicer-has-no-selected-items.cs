// Title: Select the first slicer item automatically when no selection exists before saving an Excel workbook with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set default slicer selection if none selected | How to ensure a slicer has a selected item before saving workbook using Aspose.Cells | C# code to auto-select first slicer cache item in Excel file with Aspose.Cells | Prevent empty slicer selection error Aspose.Cells .NET | Programmatically select first slicer entry when slicer is empty in Aspose.Cells
// Tags: Aspose.Cells slicer cache manipulation | C# set slicer cache item selected property | Excel slicer item selection programmatically | Aspose.Cells workbook save with slicer defaults | handle empty slicer cache Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook using Aspose.Cells, iterates through each worksheet's slicers, checks whether any slicer cache items are selected, and if none are selected it marks the first cache item as selected before saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets in the workbook
                foreach (Worksheet worksheet in workbook.Worksheets)
                {
                    // Iterate through all slicers on the current worksheet
                    foreach (Slicer slicer in worksheet.Slicers)
                    {
                        // Get the slicer cache which holds the items and their selection state
                        SlicerCache slicerCache = slicer.SlicerCache;

                        // Determine if any item is currently selected
                        bool anyItemSelected = false;
                        foreach (SlicerCacheItem item in slicerCache.SlicerCacheItems)
                        {
                            if (item.Selected) // Correct property name
                            {
                                anyItemSelected = true;
                                break;
                            }
                        }

                        // If no items are selected, select the first item in the cache
                        if (!anyItemSelected && slicerCache.SlicerCacheItems.Count > 0)
                        {
                            slicerCache.SlicerCacheItems[0].Selected = true; // Correct property name
                        }
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
