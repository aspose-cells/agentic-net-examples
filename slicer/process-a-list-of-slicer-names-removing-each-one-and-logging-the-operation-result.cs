// Title: Remove specified slicers from all worksheets in an Excel file using Aspose.Cells for .NET
// Common Searches: C# code to delete slicers by name in an Excel workbook with Aspose.Cells | How to iterate through worksheet slicer collections and remove them using Aspose.Cells .NET | Aspose.Cells example for removing multiple slicers and logging results | Delete specific slicer objects from each sheet in a .xlsx using Aspose.Cells | Saving an Excel file after slicer removal with Aspose.Cells C#
// Tags: aspnet remove slicer Aspose.Cells | worksheet slicer collection iteration .NET | bulk slicer deletion Excel Aspose | slicer removal audit log Aspose.Cells | save workbook after slicer cleanup C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;   // Required for Slicer and SlicerCollection

// The sample loads an input workbook, defines a list of slicer names, loops through every worksheet, searches each worksheet's slicer collection for matching names, removes any found slicers while outputting success or not‑found messages, handles exceptions, ensures the output directory exists, and saves the modified workbook to the specified output path.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Names of slicers to be removed
            List<string> slicerNames = new List<string>
            {
                "Slicer1",
                "Slicer2",
                "Slicer3"
            };

            // Iterate through each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Access slicer collection of the current worksheet
                SlicerCollection slicers = sheet.Slicers;

                // Process each slicer name
                foreach (string slicerName in slicerNames)
                {
                    try
                    {
                        // Find slicer by name (manual search because GetObjectByName is not available)
                        Slicer targetSlicer = null;
                        foreach (Slicer s in slicers)
                        {
                            if (s.Name.Equals(slicerName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetSlicer = s;
                                break;
                            }
                        }

                        if (targetSlicer != null)
                        {
                            // Remove the slicer
                            slicers.Remove(targetSlicer);
                            Console.WriteLine($"Removed slicer '{slicerName}' from worksheet '{sheet.Name}'.");
                        }
                        else
                        {
                            Console.WriteLine($"Slicer '{slicerName}' not found in worksheet '{sheet.Name}'.");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing slicer '{slicerName}' in worksheet '{sheet.Name}': {ex.Message}");
                    }
                }
            }

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
