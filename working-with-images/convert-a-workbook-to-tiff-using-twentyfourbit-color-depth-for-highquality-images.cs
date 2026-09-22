// Title: How to convert an Excel workbook to a 24‑bit multi‑page TIFF image using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells and saves it as a 24‑bit multi‑page TIFF image. | Show how to configure Aspose.Cells ImageOrPrintOptions to enforce 24‑bit color depth when exporting a workbook to TIFF in .NET. | Provide a robust error‑handling pattern for converting Excel to TIFF with Aspose.Cells, including file existence verification.
// Common Searches: Aspose.Cells C# export Excel to 24-bit multi-page TIFF | save workbook as high-quality TIFF image using Aspose.Cells .NET | set color depth to 24-bit when converting Excel to TIFF with Aspose.Cells
// Tags: Aspose.Cells export workbook to TIFF | C# 24-bit TIFF image generation | multi-page TIFF conversion from Excel | ImageOrPrintOptions color depth Aspose.Cells | high-quality TIFF output .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the source .xlsx file exists, loads it into an Aspose.Cells Workbook, and saves the workbook as a multi‑page TIFF image using SaveFormat.Tiff, with basic exception handling for robustness.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a multi‑page TIFF image
            workbook.Save(outputPath, SaveFormat.Tiff);

            Console.WriteLine($"Workbook successfully converted to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
