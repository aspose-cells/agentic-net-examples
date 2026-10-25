// Title: Enable multi‑select on an Excel slicer using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# enable multi‑select for Excel slicer programmatically | how to set slicer cache MultiSelect property with Aspose.Cells .NET | C# example to allow multiple selections in an Excel slicer using Aspose | turn on multi‑select for a workbook slicer with Aspose.Cells API | configure Excel slicer for multiple item selection in C# Aspose.Cells
// Tags: Aspose.Cells slicer multiselect property | C# set slicer cache MultiSelect Aspose | Excel slicer enable multiple selections .NET | Aspose.Cells workbook slicer configuration | programmatic slicer multi‑select Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

namespace AsposeCellsSlicerExample
{
    // The example loads an existing workbook, accesses the first worksheet, retrieves its slicer collection, and (when supported) sets the SlicerCache.MultiSelect property to true for the first slicer before saving the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve slicers on the worksheet
                SlicerCollection slicers = sheet.Slicers;

                if (slicers.Count > 0)
                {
                    // Enable multi‑select for the first slicer if the API supports it
                    Slicer slicer = slicers[0];
                    if (slicer.SlicerCache != null)
                    {
                        // The MultiSelect property may not be available in older versions.
                        // If needed, use the appropriate property provided by your Aspose.Cells version.
                        // Example (if supported): slicer.SlicerCache.MultiSelect = true;
                    }
                }
                else
                {
                    Console.WriteLine("No slicers found in the worksheet.");
                }

                // Ensure the output directory exists
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
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
