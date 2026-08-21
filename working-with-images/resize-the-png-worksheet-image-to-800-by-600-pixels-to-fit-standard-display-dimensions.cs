// Title: Resize an Excel worksheet to an 800×600 PNG image using Aspose.Cells for .NET
// Description: Loads an Excel workbook, selects a worksheet, configures ImageOrPrintOptions for PNG output, sets the exact pixel size to 800 × 600 (ignoring aspect ratio), renders the first page with SheetRender, saves it as "output.png", and releases resources.
// Keywords: Aspose.Cells resize worksheet image | SetDesiredSize 800 600 | export worksheet as PNG | SheetRender PNG output | ImageOrPrintOptions pixel dimensions
// Common Searches: Aspose.Cells export worksheet to PNG 800x600 | C# SetDesiredSize example Aspose.Cells | render Excel sheet as fixed size image | how to resize worksheet image with Aspose.Cells
// Developer Intent: Create a PNG snapshot of a worksheet with exact dimensions of 800 × 600 pixels.
// Use Cases: Generate uniform thumbnail previews for a web‑based workbook gallery. | Embed fixed‑size worksheet screenshots in reports, presentations, or documentation. | Provide consistent image assets for automated UI tests that require specific pixel dimensions.
// AI Prompts: Show how to keep the original aspect ratio while fitting the worksheet inside a maximum size of 800 × 600 pixels. | Provide code that batch converts every worksheet in a workbook to separate 800 × 600 PNG files using Aspose.Cells. | Explain how to adjust DPI in ImageOrPrintOptions to improve quality of the 800 × 600 PNG output.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, selects a worksheet, configures ImageOrPrintOptions for PNG output, sets the exact pixel size to 800 × 600 (ignoring aspect ratio), renders the first page with SheetRender, saves it as "output.png", and releases resources.
class ResizeWorksheetImage
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (or any specific worksheet)
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;                     // Output format PNG
        options.SetDesiredSize(800, 600, false);               // Resize to 800x600 without keeping aspect ratio

        // Create a SheetRender instance for the worksheet with the specified options
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Render the first page of the worksheet to an image file with the desired size
        sheetRender.ToImage(0, "output.png");

        // Clean up resources
        sheetRender.Dispose();

        Console.WriteLine("Worksheet image generated with size 800x600 pixels.");
    }
}
