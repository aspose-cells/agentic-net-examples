// Title: Export each worksheet from a CSV file to individual PNG images using Aspose.Cells for .NET (C#)
// Common Searches: c# convert csv workbook to png images with Aspose.Cells | how to export each sheet of a csv file as a png using Aspose.Cells | Aspose.Cells render csv worksheets to separate png files in .NET | example code for saving csv sheets as png images with Aspose.Cells C# | load csv into Aspose.Cells and generate png per worksheet
// Tags: Aspose.Cells CSV workbook PNG export | SheetRender PNG output per worksheet | ImageOrPrintOptions OnePagePerSheet usage | export multiple worksheets as separate images | C# load csv into Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads a CSV file into an Aspose.Cells Workbook, iterates through all worksheets, sets each as the active sheet, configures ImageOrPrintOptions with OnePagePerSheet, and uses SheetRender to export each worksheet as an individual PNG file named Worksheet_{index}.png. It includes file existence validation and error handling for robust execution.
class Program
{
    static void Main()
    {
        // Path to the source CSV file
        string csvPath = "input.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: The file '{csvPath}' was not found.");
            return;
        }

        try
        {
            // Load the CSV file into a workbook (Aspose.Cells treats CSV as a single‑sheet workbook)
            Workbook workbook = new Workbook(csvPath);

            // Iterate through all worksheets in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                // Get the current worksheet
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Set the current worksheet as the active sheet (required for rendering)
                workbook.Worksheets.ActiveSheetIndex = sheetIndex;

                // Configure image export options
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    // Render the entire sheet on a single page
                    OnePagePerSheet = true
                    // The default image format is PNG, so no explicit ImageFormat assignment is needed
                };

                // Create a renderer for the worksheet with the specified options
                SheetRender sheetRender = new SheetRender(sheet, imgOptions);

                // Export the first (and only) page of the sheet to a PNG file
                string outputFileName = $"Worksheet_{sheetIndex + 1}.png";

                try
                {
                    sheetRender.ToImage(0, outputFileName);
                    Console.WriteLine($"Exported: {outputFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to export sheet {sheetIndex + 1}: {ex.Message}");
                }
            }

            Console.WriteLine("Export completed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
        }
    }
}
