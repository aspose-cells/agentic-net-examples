// Title: C# – Export Excel to HTML with Page Margins and Embedded Base64 Images using Aspose.Cells
// Description: Load an Excel workbook, set left/right/top/bottom margins on the first worksheet, enable HtmlSaveOptions.ExportImagesAsBase64, and save the file as a single self‑contained HTML document.
// Keywords: Aspose.Cells HTML export | C# Excel to HTML | page margins Aspose.Cells | ExportImagesAsBase64 | embedded images HTML | HtmlSaveOptions | .NET workbook to HTML | Excel margin settings
// Common Searches: Aspose.Cells export Excel to HTML with embedded images | how to set page margins before saving workbook as HTML | C# HtmlSaveOptions ExportImagesAsBase64 example | save Excel as single HTML file Aspose.Cells | base64 image export Aspose.Cells .NET
// Developer Intent: Generate a single HTML file from an Excel workbook that preserves custom page margins and embeds all images as Base64 strings.
// Use Cases: Create a portable HTML report that mirrors the Excel layout without external image files. | Prepare HTML email content from a spreadsheet where images must be inline. | Produce printable web pages with exact margin settings for consistent pagination.
// AI Prompts: Write C# code with Aspose.Cells to load an .xlsx file, configure page margins, and save it as HTML with images embedded as Base64. | Explain the effect of HtmlSaveOptions.ExportImagesAsBase64 and how it interacts with page setup settings in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Load an Excel workbook, set left/right/top/bottom margins on the first worksheet, enable HtmlSaveOptions.ExportImagesAsBase64, and save the file as a single self‑contained HTML document.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputPath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(inputPath);

            // Enable and configure page margins for the first worksheet
            // Margins are specified in inches
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.LeftMargin = 0.5;   // Left margin
            sheet.PageSetup.RightMargin = 0.5;  // Right margin
            sheet.PageSetup.TopMargin = 0.75;   // Top margin
            sheet.PageSetup.BottomMargin = 0.75; // Bottom margin

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Embed images directly into the HTML using Base64 encoding
            htmlOptions.ExportImagesAsBase64 = true;

            // Optional: export page headers/footers if needed
            // htmlOptions.ExportPageHeaders = true;
            // htmlOptions.ExportPageFooters = true;

            // Path for the output HTML file
            string outputPath = "output.html";

            // Save the workbook as an HTML file with the specified options
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook has been exported to HTML with embedded images at: {outputPath}");
        }
    }
}
