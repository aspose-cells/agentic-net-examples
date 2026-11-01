// Title: Render an Excel worksheet to an in‑memory PNG for use as a 3D texture with Aspose.Cells in C#
// AI Prompts: Write C# code that loads a .xlsx file, renders the first worksheet to a PNG image using Aspose.Cells, and returns the PNG as a MemoryStream. | Extend the example to iterate through all pages of a worksheet and produce a separate MemoryStream for each PNG, suitable for feeding into a 3D rendering engine. | Add comprehensive error handling that checks file existence, catches rendering exceptions, and logs the size of each generated PNG stream.
// Common Searches: how to convert an Excel sheet to a PNG stream in C# with Aspose.Cells | using Aspose.Cells generate textures for Unity from Excel worksheets | render Excel worksheet to memory without saving image file C# | create PNG stream from worksheet for 3D engine using Aspose.Cells
// Tags: Aspose.Cells SheetRender to PNG MemoryStream | C# generate in‑memory worksheet image | Excel worksheet texture mapping with Aspose.Cells | render multi‑page worksheet as separate PNG streams | avoid file I/O when converting Excel to image

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads an Excel workbook, selects the first worksheet, renders it to a PNG image with Aspose.Cells' SheetRender, writes the image into a MemoryStream, and outputs the byte size, enabling the PNG to be used directly as a texture in 3D rendering pipelines.
class WorksheetTextureExample
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string excelPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file '{excelPath}' was not found.");
                return;
            }

            // Load the workbook using Aspose.Cells
            Workbook workbook = new Workbook(excelPath);

            // Choose the worksheet to render (e.g., the first worksheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure rendering options (default format is PNG)
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // One page per sheet is typical for worksheet images
                OnePagePerSheet = true
            };

            // Render the worksheet to a PNG image using SheetRender.
            // The image is stored in a memory stream to avoid intermediate files.
            using (MemoryStream pngStream = new MemoryStream())
            {
                // Create a SheetRender object for the selected worksheet
                SheetRender sheetRender = new SheetRender(sheet, renderOptions);

                // Render the first page of the worksheet (most worksheets fit on a single page)
                sheetRender.ToImage(0, pngStream);

                // Reset stream position to the beginning for any further processing
                pngStream.Position = 0;

                // For demonstration, report the size of the generated image.
                Console.WriteLine($"Worksheet rendered to PNG. Image size: {pngStream.Length} bytes.");
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
