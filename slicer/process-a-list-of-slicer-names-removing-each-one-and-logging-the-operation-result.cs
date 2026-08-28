// Title: Remove specified slicers from all worksheets in an Excel file using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads a workbook, iterates every worksheet, and deletes slicers whose names are listed in a string array, handling missing slicers without throwing. | Add console output that logs whether each named slicer was removed or not found, then save the modified workbook to a new file.
// Common Searches: how to delete slicer by name using Aspose.Cells C# | remove multiple slicers from an Excel workbook programmatically Aspose.Cells | iterate worksheets and remove slicers Aspose.Cells example | log slicer removal results in C# Aspose.Cells
// Tags: slicer removal worksheet Aspose.Cells | batch delete slicers by name C# | log slicer deletion results Aspose.Cells | save workbook after slicer removal Aspose.Cells | handle missing slicer exception Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerRemovalDemo
{
    // The program loads an existing workbook, loops through each worksheet, attempts to retrieve slicers by the names defined in an array, removes any found slicers while logging success or not‑found messages, and finally saves the updated workbook.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputWorkbook.xlsx");

            // Define the slicer names to be removed
            string[] slicerNamesToRemove = { "Slicer1", "Slicer2", "Slicer3" };

            // Iterate through each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Get the slicer collection for the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Process each slicer name
                foreach (string slicerName in slicerNamesToRemove)
                {
                    try
                    {
                        // Attempt to retrieve the slicer by its name
                        Slicer slicer = slicers[slicerName];

                        if (slicer != null)
                        {
                            // Remove the slicer from the collection
                            slicers.Remove(slicer);
                            Console.WriteLine($"Removed slicer '{slicerName}' from worksheet '{sheet.Name}'.");
                        }
                        else
                        {
                            Console.WriteLine($"Slicer '{slicerName}' not found in worksheet '{sheet.Name}'.");
                        }
                    }
                    catch (Exception ex)
                    {
                        // Indexer throws if the slicer name does not exist
                        Console.WriteLine($"Slicer '{slicerName}' not found in worksheet '{sheet.Name}'. Exception: {ex.Message}");
                    }
                }
            }

            // Save the modified workbook (replace with your desired output path)
            workbook.Save("OutputWorkbook.xlsx");
        }
    }
}
