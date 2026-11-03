// Title: Load an Excel workbook from a MemoryStream and export each worksheet as a 300 DPI PNG image using Aspose.Cells for .NET
// Common Searches: c# aspose.cells load workbook from memory stream and render to png | how to set 300 dpi when converting Excel sheets to png with aspose.cells | export each worksheet to separate high‑resolution png files using aspose.cells | sheetrender png output with custom dpi in .net
// Tags: memory stream workbook loading Aspose.Cells | worksheet to PNG conversion 300 DPI | SheetRender export separate PNG per sheet | ImageOrPrintOptions DPI configuration C# | high resolution Excel to image Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example reads an Excel file into a byte array, loads it into a Workbook via a MemoryStream, sets ImageOrPrintOptions to 300 dpi, and uses SheetRender to save each worksheet as an individual high‑resolution PNG image.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook bytes from the file
            byte[] workbookBytes = File.ReadAllBytes(inputPath);

            // Load workbook from memory stream
            using (MemoryStream ms = new MemoryStream(workbookBytes))
            {
                Workbook workbook = new Workbook(ms);

                // Configure image options: 300 DPI resolution (PNG format will be inferred from file extension)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 300,
                    VerticalResolution = 300
                };

                // Render each worksheet to a separate PNG file
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    string outputFile = $"Sheet{i + 1}.png";
                    SheetRender renderer = new SheetRender(workbook.Worksheets[i], imgOptions);
                    renderer.ToImage(0, outputFile);
                }
            }

            Console.WriteLine("Conversion completed successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
