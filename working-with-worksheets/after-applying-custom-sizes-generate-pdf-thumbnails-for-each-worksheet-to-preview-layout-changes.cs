using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply custom page setup to each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.PaperSize = PaperSizeType.PaperA4;   // set paper size
            pageSetup.FitToPagesWide = 1;                 // fit to one page wide
            pageSetup.FitToPagesTall = 1;                 // fit to one page tall
            pageSetup.LeftMargin = 0.5;                   // optional margins (in inches)
            pageSetup.RightMargin = 0.5;
            pageSetup.TopMargin = 0.5;
            pageSetup.BottomMargin = 0.5;
        }

        // Directory to store thumbnail images
        string thumbnailFolder = "Thumbnails";
        Directory.CreateDirectory(thumbnailFolder);

        // Generate a PNG thumbnail for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Configure image options
            ImageOrPrintOptions options = new ImageOrPrintOptions();
            options.ImageType = ImageType.Png;   // output as PNG
            options.OnePagePerSheet = true;      // render as a single page

            // Create SheetRender (constructor rule)
            SheetRender render = new SheetRender(sheet, options);

            // Render the first (and only) page to a PNG file (ToImage overload rule)
            string thumbnailPath = Path.Combine(thumbnailFolder, $"{sheet.Name}_thumb.png");
            render.ToImage(0, thumbnailPath);

            // Release resources
            render.Dispose();
        }

        Console.WriteLine("PDF thumbnails (PNG previews) generated successfully.");
    }
}