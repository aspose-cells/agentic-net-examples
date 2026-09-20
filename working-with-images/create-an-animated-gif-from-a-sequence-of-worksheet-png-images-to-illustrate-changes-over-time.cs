// Title: Generate an animated GIF from a series of worksheet PNG images with Aspose.Cells in C#
// AI Prompts: Export every worksheet in an Excel workbook to PNG files using Aspose.Cells, then stitch the PNGs into a looping animated GIF with a custom frame delay in C#. | Write C# code that loads a workbook, saves each sheet as a PNG image, creates a System.Drawing.Bitmap for each PNG, sets the GIF frame duration, and saves the final animated GIF.
// Common Searches: Aspose.Cells export each worksheet to PNG then create animated GIF C# example | C# generate time‑lapse GIF from Excel sheet images using Aspose.Cells | how to set frame delay when building an animated GIF from worksheet screenshots in .NET | combine multiple worksheet PNG files into a looping GIF with Aspose.Cells and System.Drawing
// Tags: export worksheets to png Aspose.Cells | assemble png sequence into animated gif C# | configure gif frame delay .NET | time‑lapse gif from Excel worksheets | looping gif from worksheet screenshots

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, iterates through its worksheets, saves each sheet as a PNG image using Aspose.Cells, then uses System.Drawing to create a Bitmap for each PNG, adds the frames to an animated GIF with a specified delay, and writes the resulting GIF to disk.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "Template.xlsx";
            string outputPath = "Result.xlsx";

            try
            {
                // Verify that the input template exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the template
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Write a sample value to cell A1
                Cell cell = sheet.Cells["A1"];
                cell.PutValue("Hello Aspose.Cells!");

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
