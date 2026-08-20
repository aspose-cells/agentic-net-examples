// Title: C# – Generate PNG thumbnails for each worksheet after custom page‑size setup with Aspose.Cells
// Description: Loads an Excel workbook, applies custom page‑setup settings (A4 paper size, fit‑to‑width, optional print area) to every worksheet, saves the modified file, and creates a single‑page PNG thumbnail for each sheet using SheetRender with OnePagePerSheet enabled.
// Keywords: Aspose.Cells | C# worksheet thumbnail | SheetRender PNG | OnePagePerSheet | custom page size | A4 paper size | Excel preview image | generate worksheet preview | ImageOrPrintOptions | export worksheet as image
// Common Searches: Aspose.Cells create worksheet thumbnail C# | How to set custom page size and render PNG preview with Aspose.Cells | Generate single‑page image for each Excel sheet using SheetRender | C# code to export Excel worksheets as PNG thumbnails | Preview Excel print layout as image Aspose.Cells
// Developer Intent: Produce PNG thumbnail images for every worksheet after configuring custom page‑setup settings.
// Use Cases: Display printable previews of Excel sheets in a web portal after applying A4 layout | Create a thumbnail gallery for report selection in a dashboard | Automate generation of preview images for batch‑processed workbooks before PDF conversion | Validate page‑setup changes by comparing generated thumbnails
// AI Prompts: Write a C# function that takes a workbook path, applies A4 paper size with fit‑to‑width, and returns file paths of PNG thumbnails for each sheet using Aspose.Cells. | Show how to change ImageOrPrintOptions to output JPEG thumbnails at 150 DPI with Aspose.Cells. | Explain how to render only the first page of each worksheet as a thumbnail while preserving custom page settings. | Provide a complete example that saves the modified workbook and generates thumbnails in a specified output folder.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, applies custom page‑setup settings (A4 paper size, fit‑to‑width, optional print area) to every worksheet, saves the modified file, and creates a single‑page PNG thumbnail for each sheet using SheetRender with OnePagePerSheet enabled.
class WorksheetThumbnailGenerator
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Apply custom page sizes and layout settings to each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Example customizations
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;   // Set paper size to A4
            sheet.PageSetup.FitToPagesWide = 1;                // Fit to one page wide
            sheet.PageSetup.FitToPagesTall = 0;                // No limit on page height
            sheet.PageSetup.PrintArea = "A1:Z100";              // Define a print area (optional)
        }

        // Optionally save the workbook after applying the custom sizes
        workbook.Save("modified.xlsx");

        // Prepare image options for rendering thumbnails
        ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
        {
            ImageType = ImageType.Png,      // Output format
            OnePagePerSheet = true          // Render each sheet as a single page
        };

        // Generate a PNG thumbnail for each worksheet
        int sheetIndex = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Create a SheetRender for the current worksheet
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // Render the first (and only) page to a PNG file
            string thumbnailFile = $"Sheet_{sheetIndex}_thumb.png";
            sheetRender.ToImage(0, thumbnailFile);

            // Release resources
            sheetRender.Dispose();

            sheetIndex++;
        }

        Console.WriteLine("Worksheet thumbnails have been generated successfully.");
    }
}
