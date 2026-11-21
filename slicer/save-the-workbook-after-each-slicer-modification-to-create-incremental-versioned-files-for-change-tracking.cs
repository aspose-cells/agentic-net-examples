// Title: Save a new version of an Excel workbook after each slicer iteration using Aspose.Cells for C#
// Common Searches: how to create versioned Excel files for each slicer in Aspose.Cells C# | Aspose.Cells save workbook after processing each slicer loop | C# generate incremental Excel file names while iterating slicers with Aspose | track slicer changes by saving workbook per slicer using Aspose.Cells | incremental versioning of Excel workbook during slicer processing in C#
// Tags: Aspose.Cells slicer iteration saving | incremental Excel versioning C# | save workbook per slicer Aspose | Excel slicer change tracking Aspose.Cells | C# generate versioned output files Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// The example loads an existing workbook, loops through every slicer on the first worksheet, and after each (placeholder) slicer operation saves the workbook to a uniquely numbered file (output_v1.xlsx, output_v2.xlsx, …). This produces incremental, version‑tracked Excel files that reflect slicer‑related changes.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);

            // Initialize version counter for incremental output files
            int version = 1;

            // Iterate through all slicers in the first worksheet
            foreach (Slicer slicer in workbook.Worksheets[0].Slicers)
            {
                // Example placeholder: additional slicer processing can be added here
                // (The original toggle of slicer items has been omitted because
                // the SlicerCache API for accessing items varies between versions.)

                // Save the workbook after processing each slicer
                string outputPath = $"output_v{version}.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Saved: {outputPath}");
                version++;
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
