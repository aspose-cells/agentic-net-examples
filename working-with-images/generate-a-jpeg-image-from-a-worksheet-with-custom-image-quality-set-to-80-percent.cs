// Title: Export an Excel worksheet to JPEG with 80% quality using Aspose.Cells (C#)
// Description: Creates a workbook, adds sample data, configures ImageOrPrintOptions to Jpeg with Quality = 80, and renders the first worksheet page to a JPEG file via SheetRender.
// Keywords: Aspose.Cells | C# | ImageOrPrintOptions | JPEG quality | export worksheet to image | SheetRender | Excel to JPEG | custom image compression | save worksheet as jpg
// Common Searches: Aspose.Cells export worksheet to JPEG | C# set JPEG quality Aspose.Cells | ImageOrPrintOptions Quality property example | render Excel sheet as JPEG with specific compression | save Excel as JPEG with 80% quality
// Developer Intent: Generate a JPEG image of a worksheet with a defined 80 % compression level.
// Use Cases: Create thumbnail previews of Excel sheets for web portals while controlling file size. | Embed worksheet images in reports where visual fidelity must be preserved. | Batch‑convert multiple worksheets to JPEG files with a consistent quality setting. | Produce email‑friendly images of spreadsheet data for quick sharing.
// AI Prompts: Show code to export all worksheets to separate JPEG files at 80% quality. | Demonstrate how to adjust JPEG quality per worksheet based on its content. | Explain how to obtain the rendered JPEG as a MemoryStream instead of writing to disk while keeping quality at 80%. | Provide tips for further reducing JPEG file size using Aspose.Cells image options.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, configures ImageOrPrintOptions to Jpeg with Quality = 80, and renders the first worksheet page to a JPEG file via SheetRender.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data to the worksheet
        worksheet.Cells["A1"].PutValue("Sample Text");
        worksheet.Cells["B2"].PutValue(123.45);
        worksheet.Cells["C3"].PutValue(DateTime.Now);

        // Configure image rendering options
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Jpeg;   // Set output format to JPEG
        options.Quality = 80;                // Set JPEG quality to 80%

        // Render the first page of the worksheet to a JPEG file
        SheetRender renderer = new SheetRender(worksheet, options);
        renderer.ToImage(0, "WorksheetImage_Quality80.jpg");

        Console.WriteLine("Worksheet rendered to JPEG with 80% quality.");
    }
}
