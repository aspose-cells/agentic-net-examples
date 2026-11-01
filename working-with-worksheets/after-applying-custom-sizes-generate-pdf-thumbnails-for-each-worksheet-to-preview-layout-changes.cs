// Title: Generate PNG thumbnails for each worksheet after applying custom column widths and row heights using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# generate PNG thumbnail for each Excel worksheet after resizing columns | How to set column width and row height then export sheet image with Aspose.Cells | Create worksheet preview images in .NET after custom column sizing | Save Excel workbook as PDF and also produce PNG previews of each sheet using Aspose.Cells
// Tags: custom column width row height Aspose.Cells | worksheet to PNG thumbnail rendering | OnePagePerSheet image options Aspose.Cells | save workbook as PDF Aspose.Cells | sanitize worksheet name for file path C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, sets column A to 20 characters and row 1 to 30 points on every worksheet, creates a thumbnails folder, and renders each sheet to a 96‑DPI PNG image using ImageOrPrintOptions with OnePagePerSheet enabled. It also saves the whole workbook as a PDF and includes a helper to make worksheet names safe for file paths.
class WorksheetPdfThumbnailGenerator
{
    static void Main()
    {
        try
        {
            // Input workbook path – verify existence to avoid FileNotFoundException
            string workbookPath = @"C:\Input\Sample.xlsx";
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Workbook not found: {workbookPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Apply custom column widths and row heights to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Set column A width to 20 characters
                sheet.Cells.SetColumnWidth(0, 20);
                // Set row 1 height to 30 points
                sheet.Cells.SetRowHeight(0, 30);
            }

            // Folder to store generated thumbnails
            string thumbnailFolder = @"C:\Output\Thumbnails";
            Directory.CreateDirectory(thumbnailFolder);

            // Iterate through each worksheet to create PNG thumbnails
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Configure image rendering options
                    ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                    {
                        // Default format is PNG; explicit setting omitted for compatibility
                        HorizontalResolution = 96,
                        VerticalResolution = 96,
                        OnePagePerSheet = true
                    };

                    // Render the worksheet to a PNG thumbnail
                    SheetRender sr = new SheetRender(sheet, imgOptions);
                    string safeSheetName = GetSafeFileName(sheet.Name);
                    string thumbnailPath = Path.Combine(thumbnailFolder, $"{safeSheetName}_Thumbnail.png");
                    sr.ToImage(0, thumbnailPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to generate thumbnail for sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Optional: Save the entire workbook as a PDF for reference
            string pdfOutputPath = @"C:\Output\Workbook.pdf";
            try
            {
                workbook.Save(pdfOutputPath, SaveFormat.Pdf);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook as PDF: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper to remove invalid filename characters
    private static string GetSafeFileName(string name)
    {
        foreach (char c in Path.GetInvalidFileNameChars())
        {
            name = name.Replace(c, '_');
        }
        return name;
    }
}
