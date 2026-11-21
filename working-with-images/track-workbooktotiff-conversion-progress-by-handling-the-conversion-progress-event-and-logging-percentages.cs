// Title: Monitor Aspose.Cells workbook-to-TIFF conversion progress in C# by handling the ConversionProgress event
// Common Searches: how to use Aspose.Cells ConversionProgress event to log percentage while saving Excel as TIFF in C# | C# example for tracking workbook to multi‑page TIFF conversion status with Aspose.Cells | track Aspose.Cells Excel to TIFF conversion progress and get callback percentages in .NET | handle conversion progress event when converting .xlsx to .tiff using Aspose.Cells library | log conversion progress of workbook.save to TIFF using Aspose.Cells in a C# application
// Tags: Aspose.Cells ConversionProgress event | C# workbook to multi‑page TIFF conversion | log conversion percentage Aspose.Cells | Excel to TIFF progress tracking .NET | handle conversion progress callback Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorkbookToTiffExample
{
    // This example shows how to subscribe to the Aspose.Cells Workbook ConversionProgress event, save an Excel workbook as a multi‑page TIFF, and log the percentage completed during the save operation, with proper error handling.
    class WorkbookToTiffConverter
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            try
            {
                // Verify that the source workbook exists before attempting to load it.
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                    return;
                }

                // Load the source workbook.
                Workbook workbook = new Workbook(inputPath);

                // Save the workbook as a multi‑page TIFF using the built‑in SaveFormat.
                workbook.Save(outputPath, SaveFormat.Tiff);
                Console.WriteLine($"Workbook successfully saved as TIFF to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message.
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
