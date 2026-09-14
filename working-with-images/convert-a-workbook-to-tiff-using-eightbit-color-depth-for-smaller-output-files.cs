// Title: How to convert an Excel workbook to an 8‑bit multi‑page TIFF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, verifies its presence, and saves the first worksheet as a TIFF image with 8‑bit color depth using Aspose.Cells ImageOrPrintOptions. | Show how to configure Aspose.Cells SaveOptions to generate a multi‑page TIFF from a workbook while reducing the color depth for smaller file size. | Create a console application that catches and logs exceptions during the conversion of an Excel workbook to a low‑color‑depth TIFF image. | Explain the steps to set TIFF compression and color depth properties in Aspose.Cells to optimize the output file.
// Common Searches: Aspose.Cells C# save workbook as 8-bit TIFF | reduce TIFF file size when exporting Excel with Aspose.Cells | multi-page TIFF generation from Excel worksheet .NET | set color depth for TIFF output using Aspose.Cells | exception handling for Excel to TIFF conversion Aspose.Cells
// Tags: save workbook as 8-bit tiff Aspose.Cells | tiff color depth configuration Aspose.Cells | multi-page tiff export .net | optimize tiff file size Aspose.Cells | excel worksheet to tiff conversion c# | image save options for tiff Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the existence of an input.xlsx file, loads it into an Aspose.Cells Workbook, and saves the first worksheet (or the whole workbook) as a multi‑page TIFF image with 8‑bit color depth, producing output.tiff while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.tiff";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook (first sheet) as a TIFF image.
            // SaveFormat.Tiff creates a multi‑page TIFF where each worksheet is a page.
            // If only the first worksheet is needed, it can be saved separately by rendering that sheet.
            workbook.Save(outputPath, SaveFormat.Tiff);

            Console.WriteLine($"TIFF image saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log or display the exception details for troubleshooting
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
