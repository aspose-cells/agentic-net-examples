// Title: Export a Worksheet to a 300 DPI Transparent PNG Using Aspose.Cells for .NET
// Description: Loads an Excel file, picks a worksheet, sets ImageOrPrintOptions to PNG with 300 DPI horizontal and vertical resolution and a transparent canvas, then renders the first page via SheetRender and saves it as an image.
// Keywords: Aspose.Cells | C# | .NET | export worksheet image | high‑resolution PNG | 300 DPI | transparent background | SheetRender | ImageOrPrintOptions | Excel to image conversion
// Common Searches: Aspose.Cells export worksheet as PNG with transparency | how to set 300 DPI for Excel image in C# | generate transparent PNG from Excel sheet using Aspose | SheetRender 300 DPI output example | C# code to save worksheet as high‑resolution image
// Developer Intent: Create a PNG file of a selected worksheet at 300 DPI while keeping the background transparent.
// Use Cases: Embedding crisp, background‑free worksheet graphics in web pages or presentations. | Producing print‑ready assets that require exact 300 DPI resolution for marketing collateral. | Generating UI thumbnails or overlays where the Excel sheet must blend seamlessly with other visuals.
// AI Prompts: Write C# code that exports the second worksheet of a workbook to a 300 DPI PNG with a transparent canvas using Aspose.Cells. | Show how to loop through all pages of a worksheet and save each as an individual 300 DPI transparent PNG. | Demonstrate switching the output format to TIFF while preserving 300 DPI resolution and a transparent background in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an Excel file, picks a worksheet, sets ImageOrPrintOptions to PNG with 300 DPI horizontal and vertical resolution and a transparent canvas, then renders the first page via SheetRender and saves it as an image.
class ExportWorksheetToImage
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Select the worksheet you want to export (by index or name)
        Worksheet worksheet = workbook.Worksheets[0]; // first worksheet

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,          // PNG supports transparency
            HorizontalResolution = 300,         // 300 DPI horizontally
            VerticalResolution = 300,           // 300 DPI vertically
            Transparent = true                  // Enable transparent background
        };

        // Create a SheetRender instance for the selected worksheet
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Export the first page (page index 0) to an image file
        string outputImagePath = "worksheet_300dpi_transparent.png";
        sheetRender.ToImage(0, outputImagePath);

        // Release resources used by SheetRender
        sheetRender.Dispose();

        Console.WriteLine($"Worksheet exported successfully to '{outputImagePath}' with 300 DPI and transparent background.");
    }
}
