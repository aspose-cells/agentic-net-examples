// Title: C# – Iterate all slicers in an Excel workbook, log selected items, clear selections, and refresh using Aspose.Cells
// AI Prompts: Write a C# program with Aspose.Cells that loads an Excel file, loops through every slicer on each worksheet, prints the values of items where Selected is true, sets Selected = false for all items, and calls Refresh on the slicer. | Generate .NET code that enumerates the SlicerCollection of a workbook, logs each selected SlicerCacheItem to the console, clears all selections, and updates the slicer view using Aspose.Cells. | Create a C# snippet that opens a workbook, accesses slicer caches, outputs selected slicer entries, deselects every entry, and refreshes the slicer to apply changes.
// Common Searches: Aspose.Cells C# list selected slicer items in a workbook | How to programmatically clear slicer selections with Aspose.Cells .NET | Refresh slicer after changing selection using Aspose.Cells API | Iterate over slicer collections on all worksheets with Aspose.Cells | Log slicer cache selected values and deselect them in C#
// Tags: Aspose.Cells iterate slicer cache items | C# clear slicer selections Aspose.Cells | Aspose.Cells refresh slicer after deselection | enumerate worksheet slicers .NET | log selected slicer items C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerDemo
{
    // // Loads an Excel workbook, iterates each worksheet's slicers, writes selected slicer item values to the console, deselects all items, and refreshes each slicer.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the workbook file. Adjust as needed.
            const string workbookPath = "input.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook if the file exists; otherwise create a new one.
                if (File.Exists(workbookPath))
                {
                    workbook = new Workbook(workbookPath);
                }
                else
                {
                    Console.WriteLine($"File '{workbookPath}' not found. Creating a new workbook.");
                    workbook = new Workbook(); // empty workbook
                }

                // Iterate through all worksheets.
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Get slicers collection for the current worksheet.
                    SlicerCollection slicers = ws.Slicers;

                    // Process each slicer.
                    for (int s = 0; s < slicers.Count; s++)
                    {
                        Slicer slicer = slicers[s];

                        // Log currently selected items.
                        foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
                        {
                            if (item.Selected)
                            {
                                Console.WriteLine($"Worksheet: {ws.Name}, Slicer: {slicer.Name}, Selected Item: {item.Value}");
                            }
                        }

                        // Deselect all items.
                        foreach (SlicerCacheItem item in slicer.SlicerCache.SlicerCacheItems)
                        {
                            item.Selected = false;
                        }

                        // Refresh slicer to apply changes.
                        slicer.Refresh();
                    }
                }

                // Optionally save the workbook if modifications were made.
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
