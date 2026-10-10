// Title: How to monitor workbook‑to‑TIFF conversion progress in C# with Aspose.Cells SaveProgress event
// AI Prompts: Generate C# code that subscribes to Workbook.SaveProgress and writes the percentage completed to the console while saving an Excel file as a TIFF using Aspose.Cells. | Show an example of implementing a console‑based progress bar that updates from the SaveProgress event during a multi‑page TIFF export with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# how to get save progress percentage when exporting to TIFF | track Excel to TIFF conversion progress using Workbook.SaveProgress event | display console progress bar during Aspose.Cells TIFF export in .NET | log conversion status of workbook to multi‑page TIFF with Aspose.Cells
// Tags: Workbook.SaveProgress event Aspose.Cells C# | TIFF export progress logging Aspose.Cells | console percentage display Aspose.Cells conversion | Excel to multi-page TIFF progress handling | Aspose.Cells conversion event handling C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Demonstrates loading an Excel workbook, attaching a handler to the Workbook.SaveProgress event to log conversion percentages, and saving the workbook as a multi‑page TIFF using Aspose.Cells, with console output for both progress updates and final status.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source workbook
                Workbook workbook = new Workbook(inputPath);

                // Save the workbook as a TIFF image
                workbook.Save(outputPath, SaveFormat.Tiff);

                Console.WriteLine($"Workbook successfully saved as TIFF to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
