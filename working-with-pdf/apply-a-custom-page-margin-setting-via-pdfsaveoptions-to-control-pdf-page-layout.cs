// Title: Apply custom worksheet margins and export an Excel workbook to PDF using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Generate C# code that sets left, right, top, and bottom margins on a worksheet, selects A4 paper size, and saves the workbook as a PDF with PdfSaveOptions preserving pagination. | Show how to configure Aspose.Cells PdfSaveOptions to turn off OnePagePerSheet while applying specific margin measurements before PDF conversion. | Provide a complete example that verifies the source .xlsx file, applies margin values in points, and writes the resulting PDF to a target path.
// Common Searches: Aspose.Cells set worksheet page margins before converting to PDF in C# | C# how to export Excel to PDF with custom margins using PdfSaveOptions | PdfSaveOptions OnePagePerSheet false margin settings Aspose.Cells example | Set A4 paper size and custom margins for PDF output with Aspose.Cells
// Tags: worksheet page margin configuration Aspose.Cells | PdfSaveOptions margin control C# | Excel to PDF conversion with A4 layout | disable pagination per sheet Aspose.Cells PDF | apply page setup margins prior to PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads input.xlsx, sets 0.5‑inch left/right and 1‑inch top/bottom margins on the first worksheet, selects A4 paper size, configures PdfSaveOptions with OnePagePerSheet = false, and saves the workbook as output.pdf.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed).
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom page margins (values are in points; 1 inch = 72 points).
            // Ensure margins are reasonable so their sum does not exceed the page width.
            sheet.PageSetup.LeftMargin = 0.5 * 72;   // 0.5 inch
            sheet.PageSetup.RightMargin = 0.5 * 72;  // 0.5 inch
            sheet.PageSetup.TopMargin = 1.0 * 72;    // 1 inch
            sheet.PageSetup.BottomMargin = 1.0 * 72; // 1 inch

            // Optionally set a standard paper size to keep margins within bounds.
            sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;

            // Create PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep original pagination based on margins.
                OnePagePerSheet = false
            };

            // Save the workbook as a PDF with the custom margins applied.
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
