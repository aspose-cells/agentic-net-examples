// Title: Delete a list of slicers from an Excel workbook and log each operation using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, iterates over a predefined list of slicer names, removes each existing slicer from the first worksheet, writes a console message indicating whether the slicer was removed or not found, and saves the modified workbook. | Create a .NET snippet that checks a worksheet's SlicerCollection for specific slicer names, deletes them if present, logs the outcome for each slicer, and includes error handling for missing input files.
// Common Searches: Aspose.Cells C# remove slicer by name from worksheet | how to programmatically delete multiple slicers in an Excel file using Aspose | C# example for iterating SlicerCollection and removing slicers | log slicer removal status when processing Excel workbook with Aspose.Cells | remove slicers from first sheet and save workbook Aspose.Cells .NET
// Tags: Aspose.Cells delete slicer C# | SlicerCollection remove slicer .NET | log slicer removal Aspose.Cells | process Excel slicers programmatically | save workbook after slicer deletion Aspose

using Aspose.Cells;
using Aspose.Cells.Slicers;
using System;
using System.Collections.Generic;
using System.IO;

// // Loads 'input.xlsx', attempts to delete slicers named Slicer1, Slicer2, and Slicer3 from the first worksheet using Aspose.Cells, writes a console message for each removal or missing slicer, and saves the updated file as 'output.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Define slicer names to be removed
            List<string> slicerNames = new List<string> { "Slicer1", "Slicer2", "Slicer3" };

            // Assume slicers are on the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            SlicerCollection slicers = worksheet.Slicers;

            foreach (string slicerName in slicerNames)
            {
                // Retrieve the slicer by name; returns null if not found
                Slicer slicer = slicers[slicerName];

                if (slicer != null)
                {
                    // Remove the slicer from the worksheet
                    slicers.Remove(slicer);
                    Console.WriteLine($"Removed slicer '{slicerName}'.");
                }
                else
                {
                    Console.WriteLine($"Slicer '{slicerName}' not found.");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
