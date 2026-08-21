// Title: Export a worksheet range to a single‑page PDF with custom margins and landscape orientation using Aspose.Cells for .NET
// Description: Creates a workbook, defines a print area (e.g., A5:D15), sets top, bottom, left and right margins, switches to landscape mode, and saves only that range as a one‑page PDF with Aspose.Cells.
// Keywords: Aspose.Cells export range PDF | custom margins PDF .NET | landscape orientation Aspose.Cells | print area to PDF | OnePagePerSheet | AllColumnsInOnePagePerSheet | PdfSaveOptions C# | page setup Aspose.Cells | Aspose.Cells PDF conversion
// Common Searches: how to export a specific cell range to PDF with Aspose.Cells | set custom margins for PDF export in Aspose.Cells C# | save worksheet as landscape PDF using Aspose.Cells | export only the print area to a single PDF page Aspose.Cells | fit all columns on one PDF page Aspose.Cells .NET
// Developer Intent: Generate a PDF that contains only the defined range, applying user‑specified margins and landscape page orientation.
// Use Cases: Produce a printable PDF of a report section with exact margin control for corporate branding. | Create landscape‑oriented PDF snapshots of data blocks for inclusion in slide decks or manuals. | Export a single‑page PDF of a print area while ensuring all columns are compressed onto one page.
// AI Prompts: Show how to export multiple non‑contiguous ranges to separate PDF pages, each with its own margin settings, using Aspose.Cells. | Provide code that adds a header and footer to the PDF while keeping custom margins and landscape orientation. | Explain how to convert margin values from centimeters to inches before applying them in PageSetup.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsRangeToPdf
{
    // Creates a workbook, defines a print area (e.g., A5:D15), sets top, bottom, left and right margins, switches to landscape mode, and saves only that range as a one‑page PDF with Aspose.Cells.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (A1:D20)
            for (int row = 0; row < 20; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Define the range to be exported (e.g., A5:D15)
            string exportRange = "A5:D15";

            // Apply page setup settings
            PageSetup pageSetup = sheet.PageSetup;

            // Set custom margins (in inches)
            pageSetup.TopMarginInch = 0.5f;      // 0.5 inch top margin
            pageSetup.BottomMarginInch = 0.5f;   // 0.5 inch bottom margin
            pageSetup.LeftMarginInch = 0.75f;    // 0.75 inch left margin
            pageSetup.RightMarginInch = 0.75f;   // 0.75 inch right margin

            // Set page orientation (Landscape)
            pageSetup.Orientation = PageOrientationType.Landscape;

            // Set the print area to the desired range
            pageSetup.PrintArea = exportRange;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the defined print area is the only content saved
                OnePagePerSheet = true,
                // Fit the defined range onto a single PDF page
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook as PDF; only the specified range will appear
            workbook.Save("ExportedRange.pdf", pdfOptions);

            Console.WriteLine("Range exported to PDF with custom margins and orientation.");
        }
    }
}
