// Title: How to programmatically select specific slicer items in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# select slicer items by value list | Set slicer selections programmatically in Excel with Aspose.Cells .NET | Clear existing slicer selections and apply new ones using Aspose.Cells | Refresh pivot tables after updating slicer cache items in C# | Load workbook, modify slicer cache items, and save with Aspose.Cells | Select multiple slicer items from an array in Aspose.Cells C#
// Tags: Aspose.Cells slicer cache manipulation | C# programmatic slicer selection | Excel slicer item selection using Aspose.Cells | Refresh pivot tables after slicer changes Aspose.Cells | Load and save workbook with updated slicer items C# | Clear slicer selections Aspose.Cells API

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;
using Aspose.Cells.Pivot;

namespace SlicerItemSelector
{
    // The example loads an existing Excel file, refreshes all pivot tables and slicers, iterates through each worksheet's slicers, clears any prior selections, selects slicer items whose values match a predefined string array, refreshes the workbook again to apply the changes, and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths to the input and output Excel files.
            // Adjust these paths as needed for your environment.
            string inputPath = "input.xlsx";
            string outputPath = "output_selected_slicer_items.xlsx";

            // Predefined list of slicer item values that should be selected.
            // Modify this array to match the values you want to select.
            string[] valuesToSelect = new string[] { "Apple", "Banana", "Cherry" };

            // Verify that the input file exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook that already contains a slicer (or slicers).
                Workbook workbook = new Workbook(inputPath);

                // Refresh all pivot tables and slicers in the workbook.
                workbook.RefreshAll();

                // Iterate through all worksheets.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // The Slicers collection holds all slicer objects placed on the worksheet.
                    SlicerCollection slicers = sheet.Slicers;

                    // Process each slicer individually.
                    for (int i = 0; i < slicers.Count; i++)
                    {
                        Slicer slicer = slicers[i];
                        SlicerCache cache = slicer.SlicerCache;

                        // Access the collection of all items in the slicer cache.
                        SlicerCacheItemCollection cacheItems = cache.SlicerCacheItems;

                        // First, clear any existing selections.
                        foreach (SlicerCacheItem item in cacheItems)
                        {
                            item.Selected = false;
                        }

                        // Then, select only the items whose Value matches the predefined list.
                        foreach (SlicerCacheItem item in cacheItems)
                        {
                            if (Array.Exists(valuesToSelect, v => string.Equals(v, item.Value, StringComparison.OrdinalIgnoreCase)))
                            {
                                item.Selected = true;
                            }
                        }

                        Console.WriteLine($"Processed slicer '{slicer.Name}' on worksheet '{sheet.Name}'.");
                    }
                }

                // Refresh again to ensure pivot tables reflect the new slicer selections.
                workbook.RefreshAll();

                // Save the modified workbook.
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Output any errors that occur during processing.
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
