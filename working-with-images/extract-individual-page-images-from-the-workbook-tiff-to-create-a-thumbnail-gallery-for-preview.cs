// Title: Create 150 dpi PNG thumbnails for each worksheet by converting an Excel workbook to a multi‑page TIFF with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file, saves it as a multi‑page TIFF, and then uses SheetRender with ImageOrPrintOptions to generate a 150 dpi PNG thumbnail for every worksheet. | Provide a method that extracts each page from a workbook‑saved TIFF and saves it as an individual PNG preview using Aspose.Cells. | Show how to change the DPI and output format when rendering Excel worksheets to PNG thumbnails after converting the workbook to TIFF with Aspose.Cells.
// Common Searches: C# Aspose.Cells generate PNG thumbnail for each Excel sheet | how to extract pages from a multi‑page TIFF created from Excel | set DPI for worksheet image rendering with Aspose.Cells | create thumbnail gallery of Excel worksheets in .NET | convert Excel workbook to TIFF then to per‑sheet PNG using Aspose
// Tags: Aspose.Cells SheetRender PNG thumbnail | multi-page TIFF extraction C# | ImageOrPrintOptions DPI configuration | Excel worksheet preview generation | convert Excel to TIFF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, saves it as a multi‑page TIFF, and then iterates through each worksheet rendering a 150 dpi PNG thumbnail using ImageOrPrintOptions and SheetRender, with error handling for missing files and runtime exceptions.
class WorkbookTiffThumbnailGenerator
{
    static void Main()
    {
        try
        {
            // Path to the source workbook (any Excel format)
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The workbook file \"{workbookPath}\" was not found.");
                return;
            }

            // Path where the multi‑page TIFF will be saved
            string tiffPath = "workbook_pages.tiff";

            // Load the workbook
            Workbook wb;
            try
            {
                wb = new Workbook(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load workbook: {ex.Message}");
                return;
            }

            // Save the workbook as a multi‑page TIFF (one page per sheet)
            try
            {
                wb.Save(tiffPath, SaveFormat.Tiff);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save TIFF: {ex.Message}");
                return;
            }

            // Verify that the TIFF file was created
            if (!File.Exists(tiffPath))
            {
                Console.WriteLine($"Error: The TIFF file \"{tiffPath}\" was not created.");
                return;
            }

            // Desired thumbnail resolution (dots per inch)
            int thumbResolution = 150;

            // Prepare image options for PNG thumbnails
            ImageOrPrintOptions thumbOptions = new ImageOrPrintOptions
            {
                // Default image format for SheetRender is PNG, so no need to set explicitly
                HorizontalResolution = thumbResolution,
                VerticalResolution = thumbResolution,
                OnePagePerSheet = true
            };

            // Generate a thumbnail for each worksheet
            for (int i = 0; i < wb.Worksheets.Count; i++)
            {
                try
                {
                    // Render the current worksheet to a PNG image using the specified options
                    SheetRender sr = new SheetRender(wb.Worksheets[i], thumbOptions);
                    string thumbPath = $"thumb_{i}.png";
                    sr.ToImage(0, thumbPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to create thumbnail for sheet {i}: {ex.Message}");
                }
            }

            Console.WriteLine("Thumbnail gallery created successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
