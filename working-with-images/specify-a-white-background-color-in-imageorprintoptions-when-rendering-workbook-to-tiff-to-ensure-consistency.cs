using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

class RenderTiffWhiteBackground
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells["A1"].PutValue("TIFF rendering with white background");

        // Configure image rendering options for TIFF
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Tiff,   // Set output format to TIFF
            Transparent = false           // Ensure the background is solid white (default)
        };

        // Render the entire workbook to a TIFF file
        WorkbookRender renderer = new WorkbookRender(workbook, options);
        renderer.ToImage("output_white_background.tiff");

        Console.WriteLine("TIFF file generated with a white background.");
    }
}