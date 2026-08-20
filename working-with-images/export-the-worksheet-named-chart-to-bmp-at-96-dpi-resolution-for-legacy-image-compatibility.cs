// Title: Export a 'Chart' worksheet to BMP at 96 DPI with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, finds the worksheet named "Chart", sets ImageOrPrintOptions to BMP with 96 × 96 DPI, and uses SheetRender to save the first page as "ChartWorksheet.bmp".
// Keywords: Aspose.Cells | C# | .NET | BMP export | Excel worksheet to image | 96 DPI | ImageOrPrintOptions | SheetRender | legacy image format | convert Excel to BMP
// Common Searches: Aspose.Cells export worksheet to BMP | C# render Excel sheet as BMP 96 DPI | Save specific worksheet as BMP using Aspose.Cells | Set DPI when converting Excel to BMP | How to create BMP image from Excel worksheet in .NET
// Developer Intent: Create a BMP image of the "Chart" worksheet at 96 DPI using Aspose.Cells.
// Use Cases: Generate BMP assets for legacy reporting tools that require 96 DPI images. | Produce thumbnail previews of selected worksheets for file‑manager displays. | Prepare printable BMP files for older hardware that only accepts BMP format.
// AI Prompts: Write C# code with Aspose.Cells to export the worksheet named "Chart" to a 96 DPI BMP file, including error handling for missing sheets. | Show how to loop through all worksheets in a workbook and save each as a BMP image at 96 DPI using Aspose.Cells. | Explain how to modify horizontal and vertical resolution or switch to another image format when rendering a worksheet with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an Excel file, finds the worksheet named "Chart", sets ImageOrPrintOptions to BMP with 96 × 96 DPI, and uses SheetRender to save the first page as "ChartWorksheet.bmp".
class ExportWorksheetToBmp
{
    static void Main()
    {
        // Load the workbook (adjust the path as needed)
        Workbook workbook = new Workbook("input.xlsx");

        // Locate the worksheet named "Chart"
        Worksheet chartSheet = workbook.Worksheets["Chart"];
        if (chartSheet == null)
        {
            Console.WriteLine("Worksheet named 'Chart' was not found.");
            return;
        }

        // Configure image options: BMP format with 96 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Bmp,
            HorizontalResolution = 96,
            VerticalResolution = 96
        };

        // Render the first page of the worksheet to a BMP file
        SheetRender renderer = new SheetRender(chartSheet, options);
        string outputFile = "ChartWorksheet.bmp";
        renderer.ToImage(0, outputFile);

        Console.WriteLine($"Worksheet exported successfully to: {outputFile}");
    }
}
