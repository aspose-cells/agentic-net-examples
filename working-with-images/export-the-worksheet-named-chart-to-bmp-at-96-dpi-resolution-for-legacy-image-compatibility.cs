using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class ExportWorksheetToBmp
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the worksheet named "Chart"
        Worksheet chartSheet = workbook.Worksheets["Chart"];
        if (chartSheet == null)
        {
            Console.WriteLine("Worksheet named 'Chart' was not found.");
            return;
        }

        // Configure image options: BMP format, 96 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Bmp;          // Set output format to BMP
        options.VerticalResolution = 96;            // Set vertical DPI to 96
        // HorizontalResolution defaults to 96 DPI; set explicitly if needed:
        // options.HorizontalResolution = 96;

        // Render the worksheet to an image
        SheetRender renderer = new SheetRender(chartSheet, options);
        // Render the first (and usually only) page to a BMP file
        renderer.ToImage(0, "Chart.bmp");

        Console.WriteLine("Worksheet 'Chart' exported to Chart.bmp at 96 DPI.");
    }
}