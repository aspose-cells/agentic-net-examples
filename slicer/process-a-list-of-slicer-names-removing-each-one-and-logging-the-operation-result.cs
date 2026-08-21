// Title: Aspose.Cells .NET – Delete Specified Slicers from an Excel Workbook
// Description: A C# sample that loads an Excel file with Aspose.Cells, accesses the slicer collection on the first worksheet, loops through a predefined list of slicer names, removes each found slicer, logs successes or missing items, catches unexpected errors, and saves the updated workbook to a new file.
// Keywords: Aspose.Cells | C# | Excel slicer removal | delete slicer by name | slicer collection | batch slicer delete | Aspose.Cells .NET example
// Common Searches: how to delete specific slicers using Aspose.Cells for .NET | C# code to remove slicers from an Excel worksheet | Aspose.Cells remove slicer collection items | batch delete Excel slicers by name | Aspose.Cells slicer management tutorial
// Developer Intent: Programmatically eliminate selected slicers from a workbook and write the changes back to disk.
// Use Cases: Strip out development‑only slicers before publishing a workbook to end users. | Clear stale slicer filters after an automated data refresh. | Process multiple workbooks in a pipeline to purge a known set of slicers.
// AI Prompts: Generate C# code with Aspose.Cells that iterates over a list of slicer names, removes each if present, logs the outcome, and saves the workbook. | Show how to safely handle missing slicers and unexpected exceptions when deleting slicers in an Excel file using Aspose.Cells. | Provide a concise explanation of the Aspose.Cells SlicerCollection API for removing slicers by name.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace SlicerRemovalDemo
{
    // A C# sample that loads an Excel file with Aspose.Cells, accesses the slicer collection on the first worksheet, loops through a predefined list of slicer names, removes each found slicer, logs successes or missing items, catches unexpected errors, and saves the updated workbook to a new file.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that contains slicers
            string inputPath = "InputWorkbook.xlsx";
            string outputPath = "OutputWorkbook.xlsx";

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Assume slicers are on the first worksheet (adjust as needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the slicer collection for the worksheet
            SlicerCollection slicers = worksheet.Slicers;

            // List of slicer names to be removed
            List<string> slicerNamesToRemove = new List<string>
            {
                "Slicer1",
                "Slicer2",
                "Slicer3"
            };

            foreach (string slicerName in slicerNamesToRemove)
            {
                try
                {
                    // Retrieve the slicer by its name using the string indexer
                    Slicer slicer = slicers[slicerName];

                    if (slicer != null)
                    {
                        // Remove the slicer from the collection
                        slicers.Remove(slicer);
                        Console.WriteLine($"Removed slicer: {slicerName}");
                    }
                    else
                    {
                        Console.WriteLine($"Slicer not found: {slicerName}");
                    }
                }
                catch (Exception ex)
                {
                    // Log any unexpected errors during removal
                    Console.WriteLine($"Error removing slicer '{slicerName}': {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");
        }
    }
}
