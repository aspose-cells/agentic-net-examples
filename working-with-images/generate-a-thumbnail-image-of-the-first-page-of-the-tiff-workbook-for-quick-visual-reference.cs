using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Drawing;

namespace AsposeCellsThumbnailDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source workbook (can be an existing Excel file)
            string sourceWorkbookPath = "source.xlsx";

            // Path where the thumbnail image will be saved
            string thumbnailPath = "thumbnail_page0.png";

            // Load the workbook from file
            Workbook workbook = new Workbook(sourceWorkbookPath);

            // Configure rendering options for a small thumbnail
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                // Render as PNG (supports transparency and small file size)
                ImageType = ImageType.Png,

                // Render only one page per sheet to simplify thumbnail generation
                OnePagePerSheet = true,

                // Reduce resolution for a thumbnail (e.g., 96 DPI)
                HorizontalResolution = 96,
                VerticalResolution = 96,

                // Optional: set page index to start from the first page (default is 0)
                PageIndex = 0
            };

            // Create a WorkbookRender instance (required after any page setup changes)
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

            // Render the first page (page index 0) to an image file
            // This uses the rule: WorkbookRender.ToImage(int, string)
            renderer.ToImage(0, thumbnailPath);

            // Clean up resources
            renderer.Dispose();

            Console.WriteLine($"Thumbnail of the first page saved to: {thumbnailPath}");
        }
    }
}