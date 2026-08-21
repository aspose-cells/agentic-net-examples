// Title: Export an Aspose.Cells Workbook to a 150 DPI TIFF image (C#)
// Description: Demonstrates how to create a workbook, set ImageOrPrintOptions to TIFF format with 150 DPI horizontal and vertical resolution, and render the entire workbook to a TIFF file using WorkbookRender in Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | TIFF export | 150 DPI | ImageOrPrintOptions | WorkbookRender | Excel to image | set resolution
// Common Searches: Aspose.Cells render workbook to TIFF with 150 DPI | C# set horizontal and vertical DPI when exporting Excel to TIFF | How to export Excel as high‑resolution TIFF using Aspose.Cells | ImageOrPrintOptions DPI settings Aspose.Cells .NET
// Developer Intent: Generate a TIFF file from a workbook where both horizontal and vertical resolutions are fixed at 150 DPI.
// Use Cases: Produce print‑ready TIFF files from Excel data for marketing materials. | Create high‑resolution images for archival of financial reports. | Export worksheets to TIFF for inclusion in PDF portfolios while maintaining exact DPI.
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to a multi‑page TIFF at 300 DPI. | Explain how to adjust compression type and image quality when saving a TIFF with Aspose.Cells. | Show how to render selected worksheets to separate TIFF files with custom DPI settings.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Demonstrates how to create a workbook, set ImageOrPrintOptions to TIFF format with 150 DPI horizontal and vertical resolution, and render the entire workbook to a TIFF file using WorkbookRender in Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("TIFF rendering at 150 DPI");

        // Configure image options for TIFF output with 150 DPI resolution
        ImageOrPrintOptions options = new ImageOrPrintOptions();
        options.ImageType = ImageType.Tiff;          // Set output format to TIFF
        options.HorizontalResolution = 150;          // Horizontal DPI
        options.VerticalResolution = 150;            // Vertical DPI

        // Initialize the workbook renderer with the workbook and options
        WorkbookRender renderer = new WorkbookRender(workbook, options);

        // Render the entire workbook to a TIFF file
        string outputPath = "output_150dpi.tiff";
        renderer.ToImage(outputPath);

        Console.WriteLine($"Workbook successfully rendered to TIFF at 150 DPI: {outputPath}");
    }
}
