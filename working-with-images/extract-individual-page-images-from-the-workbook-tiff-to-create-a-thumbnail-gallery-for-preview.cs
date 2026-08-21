// Title: Create PNG thumbnails for each Excel worksheet page with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook, sets ImageOrPrintOptions.OnePagePerSheet, uses WorkbookRender to iterate over all pages, and saves each page as an individual PNG thumbnail in a specified folder. Works on Windows, Linux and macOS with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# thumbnail generation | Excel to PNG | WorkbookRender | OnePagePerSheet | ImageOrPrintOptions | Excel preview images | batch thumbnail creation | Aspose.Cells rendering .NET
// Common Searches: Aspose.Cells generate thumbnail per worksheet | C# render Excel sheet as PNG | How to export Excel pages to images using Aspose | Create image gallery from Excel workbook .NET | Save each Excel sheet as PNG file
// Developer Intent: Generate a PNG thumbnail for every rendered page of an Excel workbook.
// Use Cases: Build a web‑based preview gallery by converting each sheet to a thumbnail image. | Automate documentation pipelines that need page‑by‑page visual assets from reports. | Add visual regression checks in CI/CD to compare generated thumbnails against baselines. | Create email or chat attachments that show a quick preview of spreadsheet content. | Provide mobile apps with lightweight sheet previews without opening the full workbook.
// AI Prompts: Write C# code using Aspose.Cells to create 150 × 150 px PNG thumbnails for each worksheet page. | Show how to render workbook pages to a MemoryStream and return the images as Base64 strings. | Explain how to configure ImageOrPrintOptions DPI and compression for high‑quality thumbnails. | Provide a PowerShell script that calls a compiled .NET exe to batch‑process multiple workbooks.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, sets ImageOrPrintOptions.OnePagePerSheet, uses WorkbookRender to iterate over all pages, and saves each page as an individual PNG thumbnail in a specified folder. Works on Windows, Linux and macOS with Aspose.Cells for .NET.
class WorkbookThumbnailGenerator
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string workbookPath = "input.xlsx";

            // Verify that the workbook file exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: Workbook file not found at '{workbookPath}'.");
                return;
            }

            // Directory to store generated thumbnails
            string thumbnailDir = "thumbnails";
            Directory.CreateDirectory(thumbnailDir);

            // Load the workbook
            Workbook workbook;
            try
            {
                workbook = new Workbook(workbookPath);
            }
            catch (Exception exLoad)
            {
                Console.WriteLine($"Failed to load workbook: {exLoad.Message}");
                return;
            }

            // Configure rendering options: one page per sheet, default PNG output
            ImageOrPrintOptions renderOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
                // Default image format is PNG; additional options can be set here if needed
            };

            // Create a renderer for the workbook (WorkbookRender does not implement IDisposable)
            WorkbookRender renderer = new WorkbookRender(workbook, renderOptions);

            // Iterate through each rendered page and save as a thumbnail image
            for (int pageIndex = 0; pageIndex < renderer.PageCount; pageIndex++)
            {
                try
                {
                    string thumbPath = Path.Combine(thumbnailDir, $"thumb_page_{pageIndex}.png");

                    // Render the current page directly to a file
                    renderer.ToImage(pageIndex, thumbPath);

                    Console.WriteLine($"Thumbnail saved: {thumbPath}");
                }
                catch (Exception exPage)
                {
                    Console.WriteLine($"Failed to render page {pageIndex}: {exPage.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
