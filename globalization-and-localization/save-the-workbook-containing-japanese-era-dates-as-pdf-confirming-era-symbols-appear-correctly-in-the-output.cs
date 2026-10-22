// Title: Export an Excel workbook with Japanese era formatted dates to PDF using Aspose.Cells for .NET
// Common Searches: how to export Japanese era formatted dates to PDF with Aspose.Cells C# | Aspose.Cells PDF save options for preserving Japanese era symbols | C# code to embed fonts when saving Excel to PDF for Japanese locale | set custom number format ggge for Japanese era dates in Aspose.Cells | auto fit columns for Japanese date display before PDF export Aspose.Cells
// Tags: Japanese era date format ggge Aspose.Cells | PdfSaveOptions embed fonts Japanese characters | Export Excel to PDF with locale-specific date format | PdfA1b compliance for Japanese PDFs Aspose.Cells | AutoFitColumns for Japanese date display

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example creates a new Workbook, writes a DateTime value to cell A1, applies a custom Japanese era number format (ggge) with the Japanese locale, auto‑fits the column, configures PdfSaveOptions to embed fonts and use PDF/A‑1b compliance, ensures the output directory exists, and saves the workbook as JapaneseEraDates.pdf.
class JapaneseEraPdfExport
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Define a date (e.g., 2023-04-01) that will be displayed using Japanese era format
            DateTime date = new DateTime(2023, 4, 1);

            // Set the date value in cell A1
            Cell cell = sheet.Cells["A1"];
            cell.PutValue(date);

            // Apply Japanese era number format to the cell
            // The format string uses the Japanese locale and the era (ggge) pattern
            Style style = cell.GetStyle();
            style.Custom = "[$-ja-JP]ggge\"年\"m\"月\"d\"日\"";
            cell.SetStyle(style);

            // Adjust column width for better visibility
            sheet.AutoFitColumns();

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Embed all fonts to preserve Japanese characters in the PDF
                EmbedStandardWindowsFonts = true,
                // Use PDF/A-1b compliance if needed
                Compliance = PdfCompliance.PdfA1b
                // No need to set SaveFormat; it is implied by PdfSaveOptions
            };

            // Define output file path
            string outputPath = "JapaneseEraDates.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
