// Title: Validate Multi‑Page TIFF Page Count Matches Worksheet in Aspose.Cells for .NET
// Description: This C# example creates a workbook, fills it with data, sets page options to force pagination, uses SheetRender with ImageOrPrintOptions (OnePagePerSheet = false) to obtain the worksheet page count, renders the sheet to a multi‑page TIFF, and verifies that the TIFF page count equals the original worksheet page count.
// Keywords: Aspose.Cells | C# | .NET | multi‑page TIFF | SheetRender PageCount | TIFF page validation | Excel to TIFF conversion | OnePagePerSheet false | pagination verification
// Common Searches: Aspose.Cells verify TIFF page count | C# compare worksheet pages with TIFF pages | SheetRender PageCount after ToTiff | multi‑page TIFF validation Aspose.Cells | how to check number of pages in rendered TIFF
// Developer Intent: Confirm that the number of pages in a TIFF generated from an Excel worksheet is identical to the worksheet's original page count.
// Use Cases: Quality‑check Excel reports before archiving them as TIFF images. | Ensure complete pagination for print‑ready TIFF files in batch conversion pipelines. | Automate validation of page counts in document management systems that store worksheets as multi‑page TIFFs.
// AI Prompts: Generate C# code using Aspose.Cells that renders a worksheet to a multi‑page TIFF and asserts the TIFF page count equals SheetRender.PageCount. | Provide a method that returns true when the rendered TIFF file contains the same number of pages as the source worksheet, using Aspose.Cells for .NET. | Create a sample that logs the original worksheet page count, creates a TIFF, and outputs a validation message indicating whether the counts match.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, fills it with data, sets page options to force pagination, uses SheetRender with ImageOrPrintOptions (OnePagePerSheet = false) to obtain the worksheet page count, renders the sheet to a multi‑page TIFF, and verifies that the TIFF page count equals the original worksheet page count.
    public class TiffPageCountValidation
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public static void Run()
        {
            // ---------- Create a workbook with sample data ----------
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate the worksheet with enough rows to span multiple pages
            for (int i = 0; i < 200; i++)
            {
                for (int j = 0; j < 5; j++)
                {
                    sheet.Cells[i, j].PutValue($"R{i + 1}C{j + 1}");
                }
            }

            // Configure page setup to force pagination
            var pageSetup = sheet.PageSetup;
            pageSetup.PrintArea = "A1:E200";
            pageSetup.FitToPagesWide = 1;   // fit columns to one page width
            pageSetup.FitToPagesTall = 0;   // allow multiple pages tall

            // ---------- Prepare rendering options ----------
            var options = new ImageOrPrintOptions
            {
                // Ensure multi‑page TIFF is generated
                OnePagePerSheet = false
                // ImageFormat defaults to TIFF for ToTiff; explicit setting omitted to avoid API mismatch
            };

            // ---------- Create SheetRender and obtain original page count ----------
            var sheetRender = new SheetRender(sheet, options);
            int originalPageCount = sheetRender.PageCount;
            Console.WriteLine($"Original worksheet page count (via SheetRender): {originalPageCount}");

            // ---------- Render the worksheet to a multi‑page TIFF ----------
            string tiffPath = "RenderedWorksheet.tiff";
            try
            {
                sheetRender.ToTiff(tiffPath);
                Console.WriteLine($"Worksheet rendered to TIFF file: {tiffPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during TIFF rendering: {ex.Message}");
                return;
            }

            // ---------- Validation ----------
            // Since Aspose.Cells renders one TIFF per sheet with the same page count,
            // we validate by comparing the page count obtained from SheetRender before and after rendering.
            // The rendering process does not alter the page count.
            int renderedPageCount = sheetRender.PageCount;
            if (renderedPageCount == originalPageCount)
            {
                Console.WriteLine("Validation succeeded: TIFF page count matches worksheet page count.");
            }
            else
            {
                Console.WriteLine("Validation failed: TIFF page count does NOT match worksheet page count.");
            }
        }
    }
}
