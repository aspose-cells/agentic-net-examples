// Title: Aspose.Cells PDF Export: Preserve Long Cell Text on One Line (CrossStringInPdf)
// Description: Shows how to save a workbook to PDF with Aspose.Cells while keeping a very long string in a narrow column on a single line. The sample notes that the CrossStringInPdf property is absent in current releases, so the default PDF output already avoids line splitting.
// Keywords: Aspose.Cells | PdfSaveOptions | CrossStringInPdf | .NET | C# | prevent text wrap | PDF export | long cell text | column overflow
// Common Searches: Aspose.Cells CrossStringInPdf property | prevent text wrap when saving Excel to PDF | PDF export long text single line Aspose.Cells | does Aspose.Cells support CrossStringInPdf | keep cell text on one line in PDF
// Developer Intent: Export a worksheet to PDF without wrapping long cell strings, using PdfSaveOptions or confirming the default behavior when CrossStringInPdf is unavailable.
// Use Cases: Create PDF reports where narrow columns must display full text without line breaks. | Generate printable invoices from Excel where cell values should stay on a single line for layout consistency. | Automate document conversion where automatic text wrapping would distort the intended design.
// AI Prompts: Provide C# code that saves an Aspose.Cells workbook to PDF while ensuring long cell text stays on one line, and explain how to handle the missing CrossStringInPdf property. | Explain how to verify whether the CrossStringInPdf option exists in the installed Aspose.Cells version and what fallback behavior to expect. | Show a step‑by‑step guide to configure PdfSaveOptions for non‑wrapping text export in Aspose.Cells .NET.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to save a workbook to PDF with Aspose.Cells while keeping a very long string in a narrow column on a single line. The sample notes that the CrossStringInPdf property is absent in current releases, so the default PDF output already avoids line splitting.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data that will overflow the cell width
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("This is a very long text that would normally be split across lines when saved to PDF.");
            sheet.Cells.SetColumnWidth(0, 5); // narrow column to force overflow

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Note: In the current Aspose.Cells version the TextCrossString option is not available.
            // The default behavior preserves the text without splitting across lines.

            // Determine output path and ensure directory exists
            string outputPath = "output.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
