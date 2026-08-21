// Title: Generate a PNG thumbnail of the first worksheet page with Aspose.Cells for .NET (C#)
// Description: This C# example loads an Excel workbook, configures ImageOrPrintOptions for a 96 dpi PNG, creates a WorkbookRender instance, and saves the first sheet (page index 0) as a compact thumbnail for quick visual reference.
// Keywords: Aspose.Cells | C# thumbnail generation | WorkbookRender | Excel to PNG | low DPI image | first worksheet preview | image rendering API | Aspose.Cells .NET | Excel preview thumbnail | render workbook page
// Common Searches: Aspose.Cells create thumbnail of first Excel sheet | C# render worksheet to PNG using Aspose.Cells | How to generate low‑resolution preview of an Excel workbook | WorkbookRender ToImage example C# | Create file‑browser icons from Excel files
// Developer Intent: Produce a small PNG preview of the workbook’s first page using Aspose.Cells.
// Use Cases: Show document previews in a web portal or intranet | Display icons for Excel files in a desktop file manager | Include page snapshots in automated test logs | Speed up content indexing by storing lightweight images
// AI Prompts: Modify the code to create thumbnails for every worksheet in the workbook. | Change the output format to JPEG and adjust the resolution dynamically. | Add robust error handling for missing files, unsupported formats, and permission issues.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// This C# example loads an Excel workbook, configures ImageOrPrintOptions for a 96 dpi PNG, creates a WorkbookRender instance, and saves the first sheet (page index 0) as a compact thumbnail for quick visual reference.
class ThumbnailGenerator
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        string workbookPath = "input.xlsx";
        Workbook workbook = new Workbook(workbookPath);

        // Configure rendering options for a small thumbnail image
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Png;               // Output format
        options.HorizontalResolution = 96;               // Lower DPI for thumbnail
        options.VerticalResolution = 96;

        // Create a renderer for the workbook
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the first page (index 0) to a thumbnail file
        string thumbnailPath = "thumbnail_page0.png";
        renderer.ToImage(0, thumbnailPath);               // Uses WorkbookRender.ToImage(int, string)

        Console.WriteLine($"Thumbnail of first page saved to: {thumbnailPath}");
    }
}
