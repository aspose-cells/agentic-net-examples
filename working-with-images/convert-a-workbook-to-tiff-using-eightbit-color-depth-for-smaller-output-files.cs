// Title: Convert an Excel workbook to an 8‑bit TIFF image using Aspose.Cells for .NET
// Common Searches: how to export an .xlsx workbook to an 8-bit TIFF file with Aspose.Cells in C# | C# code sample for saving Excel as a low‑size TIFF image using Aspose | Aspose.Cells tutorial for converting Excel to TIFF with reduced color depth | save workbook as TIFF with smaller file size in .NET using Aspose.Cells | example of converting Excel to TIFF format with Aspose.Cells library
// Tags: Aspose.Cells TIFF image export | C# Excel to TIFF conversion | reduce TIFF file size with Aspose | TIFF export using SaveFormat.Tiff | image color depth control in Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, and saves the workbook as a TIFF image (output.tiff) using the SaveFormat.Tiff option, with basic error handling and console output.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Ensure the source workbook exists before attempting to load it.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook.
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a TIFF file using default options.
            workbook.Save(outputPath, SaveFormat.Tiff);

            Console.WriteLine($"Workbook successfully saved as TIFF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
