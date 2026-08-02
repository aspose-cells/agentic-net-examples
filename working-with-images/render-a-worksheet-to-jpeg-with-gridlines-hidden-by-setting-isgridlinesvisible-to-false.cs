// Title: Render an Excel worksheet to JPEG without gridlines using Aspose.Cells for .NET
// Description: Creates a workbook, adds sample data, disables worksheet gridlines (IsGridlinesVisible = false), configures JPEG options (ImageType, Quality), renders the first sheet page to a JPEG file with SheetRender, and releases resources.
// Keywords: Aspose.Cells | C# | .NET | render worksheet to JPEG | hide gridlines | ImageOrPrintOptions | SheetRender | export Excel as image | JPEG quality setting
// Common Searches: Aspose.Cells export worksheet to JPEG without gridlines | C# hide Excel gridlines when rendering to image | set JPEG quality in Aspose.Cells ImageOrPrintOptions | render specific sheet page to JPEG using Aspose.Cells | how to disable gridlines in Excel image export .NET
// Developer Intent: Generate a JPEG image of an Excel worksheet while keeping gridlines invisible.
// Use Cases: Display spreadsheet previews on a website without the default grid pattern. | Create high‑resolution JPEG reports from Excel data for marketing emails. | Produce thumbnail images of worksheets for document management systems.
// AI Prompts: Show C# code to render a worksheet to PNG with gridlines visible using Aspose.Cells. | Explain how to adjust DPI and background color in ImageOrPrintOptions for JPEG export. | Provide a loop that saves every worksheet in a workbook as separate JPEG files.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Creates a workbook, adds sample data, disables worksheet gridlines (IsGridlinesVisible = false), configures JPEG options (ImageType, Quality), renders the first sheet page to a JPEG file with SheetRender, and releases resources.
class RenderWorksheetToJpeg
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample data so the rendered image has content
        worksheet.Cells["A1"].PutValue("Sample Data");
        worksheet.Cells["A2"].PutValue("Gridlines are hidden in this JPEG");

        // Hide gridlines for the worksheet
        worksheet.IsGridlinesVisible = false;

        // Configure image options for JPEG output
        ImageOrPrintOptions options = new ImageOrPrintOptions
        {
            ImageType = ImageType.Jpeg, // Set output format to JPEG
            Quality = 90                // Optional: set JPEG quality (0-100)
        };

        // Create a SheetRender instance using the worksheet and image options
        SheetRender sheetRender = new SheetRender(worksheet, options);

        // Render the first page (index 0) to a JPEG file
        sheetRender.ToImage(0, "output.jpg");

        // Release resources used by the renderer
        sheetRender.Dispose();
    }
}
