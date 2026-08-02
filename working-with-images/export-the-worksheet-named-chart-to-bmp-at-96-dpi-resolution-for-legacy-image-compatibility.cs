// Title: Export the "Chart" worksheet to a 96 DPI BMP image using Aspose.Cells for .NET
// Description: Loads an Excel file, selects the worksheet named "Chart", sets ImageOrPrintOptions to BMP format with 96 dpi horizontal and vertical resolution, and renders the first page to a BMP file (ChartWorksheet.bmp) via SheetRender.
// Keywords: Aspose.Cells BMP export | C# export worksheet to bitmap | 96 DPI Excel image | SheetRender BMP Aspose | ImageOrPrintOptions resolution | export chart worksheet as BMP | .NET Excel to BMP
// Common Searches: Aspose.Cells export worksheet to BMP 96 dpi | C# render Excel sheet as bitmap image | How to save a specific worksheet as BMP using Aspose | Set DPI when converting Excel to BMP with Aspose.Cells | Export chart worksheet to BMP file .NET
// Developer Intent: Generate a BMP image of the "Chart" worksheet at a fixed 96 dpi resolution.
// Use Cases: Create legacy‑compatible BMP thumbnails of chart worksheets for reporting tools that only accept BMP files. | Produce bitmap images for printing on devices that require a standard 96 dpi resolution. | Automate batch conversion of selected worksheets to BMP with consistent DPI for archival storage.
// AI Prompts: Write C# code with Aspose.Cells that exports a worksheet named "Chart" to a BMP file at 96 dpi, including error handling for missing worksheets. | Show how to loop through all pages of a worksheet and save each page as a separate 96 dpi BMP image using Aspose.Cells. | Explain the effect of changing HorizontalResolution and VerticalResolution on BMP file size and quality when rendering an Excel worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an Excel file, selects the worksheet named "Chart", sets ImageOrPrintOptions to BMP format with 96 dpi horizontal and vertical resolution, and renders the first page to a BMP file (ChartWorksheet.bmp) via SheetRender.
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
            Console.WriteLine("Worksheet named 'Chart' not found.");
            return;
        }

        // Set image rendering options: BMP format, 96 DPI (default, set explicitly)
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Bmp,
            HorizontalResolution = 96,
            VerticalResolution = 96
        };

        // Render the worksheet to an image. If the sheet spans multiple pages,
        // render each page; here we assume a single page for simplicity.
        SheetRender sheetRender = new SheetRender(chartSheet, options);
        sheetRender.ToImage(0, "ChartWorksheet.bmp");

        Console.WriteLine("Worksheet 'Chart' exported to BMP at 96 DPI.");
    }
}
