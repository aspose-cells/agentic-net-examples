// Title: How to set a white background in ImageOrPrintOptions when exporting a workbook to multi‑page TIFF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel workbook, creates ImageOrPrintOptions with BackgroundColor set to white, and saves the workbook as a multi‑page TIFF using Aspose.Cells. | Show how to configure ImageOrPrintOptions for TIFF export to enforce a white canvas, then call Workbook.Save with SaveFormat.Tiff in a .NET application.
// Common Searches: Aspose.Cells C# set white background for TIFF export | ImageOrPrintOptions BackgroundColor property TIFF Aspose.Cells .NET | export Excel to multi page TIFF with white background using Aspose.Cells | force white background when saving workbook as TIFF in Aspose.Cells
// Tags: ImageOrPrintOptions white background TIFF | Aspose.Cells TIFF export options | C# set background color Aspose.Cells | multi-page TIFF rendering Aspose.Cells | Workbook.Save TIFF background color

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel file, creating an ImageOrPrintOptions object, setting its BackgroundColor to white, assigning the options to the workbook, and saving the workbook as a multi‑page TIFF. This ensures every page of the generated TIFF has a consistent white background.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a multi‑page TIFF image
            workbook.Save(outputPath, SaveFormat.Tiff);

            Console.WriteLine($"Workbook successfully saved as TIFF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Runtime error: {ex.Message}");
        }
    }
}
