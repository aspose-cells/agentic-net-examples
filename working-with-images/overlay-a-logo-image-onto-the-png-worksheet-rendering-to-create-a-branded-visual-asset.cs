// Title: Add a PNG logo to an Excel worksheet and export the sheet as a branded PNG image using Aspose.Cells for .NET
// AI Prompts: Insert a logo.png picture at cell A1 of the first worksheet and render the sheet to output.png with Aspose.Cells. | Place a PNG watermark in the top‑left corner of a worksheet, then generate a PNG snapshot of the sheet using ImageOrPrintOptions. | Implement graceful handling for missing logo files while adding a picture to a worksheet before exporting it as a PNG image.
// Common Searches: Aspose.Cells C# add picture to worksheet before rendering to PNG | How to embed a logo in an Excel sheet image using Aspose.Cells .NET | Render Excel worksheet as PNG with custom picture overlay Aspose.Cells | C# Aspose.Cells picture insertion error handling when file not found
// Tags: worksheet picture insertion Aspose.Cells | worksheet PNG rendering Aspose.Cells | logo overlay Excel Aspose.Cells | handling missing picture file Aspose.Cells | ImageOrPrintOptions PNG output Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads 'input.xlsx', adds 'logo.png' to the first worksheet at the top‑left corner, and renders the worksheet to 'output.png' as a PNG image using Aspose.Cells, with basic error handling for missing files.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string logoPath = "logo.png";
            const string outputPath = "output.png";

            // Verify input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Add logo picture if the file exists
            if (File.Exists(logoPath))
            {
                try
                {
                    // Adds the picture at the top‑left corner (adjust positioning as needed)
                    sheet.Pictures.Add(0, 0, logoPath);
                }
                catch (Exception picEx)
                {
                    Console.WriteLine($"Failed to add logo picture: {picEx.Message}");
                }
            }

            // Set rendering options (default format inferred from output file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();

            // Render the worksheet to a PNG file
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);
            sheetRender.ToImage(0, outputPath);

            Console.WriteLine($"Branded image saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
