// Title: Convert Aspose.Cells Workbook to Grayscale PDF (Black‑and‑White) in C#
// Description: Creates a workbook, enables the BlackAndWhite page‑setup flag for every worksheet, applies default PdfSaveOptions, and saves the file as a monochrome PDF (output_grayscale.pdf).
// Keywords: Aspose.Cells PDF conversion | grayscale PDF export | BlackAndWhite page setup | C# Excel to PDF | monochrome workbook export
// Common Searches: Aspose.Cells export Excel to grayscale PDF | set black and white printing before PDF save | C# convert workbook to monochrome PDF | how to force PDF output to black and white with Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook while forcing all pages to render in black‑and‑white.
// Use Cases: Produce cost‑effective printable reports that must be monochrome. | Create archival or regulatory documents that require black‑and‑white output. | Batch‑process multiple workbooks to ensure consistent grayscale appearance across a document set.
// AI Prompts: Write a script that scans a directory of .xlsx files and converts each to a grayscale PDF using Aspose.Cells. | Explain how to embed a custom grayscale color profile in PdfSaveOptions for Aspose.Cells. | Show code that applies the BlackAndWhite setting only to selected worksheets before exporting to PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Creates a workbook, enables the BlackAndWhite page‑setup flag for every worksheet, applies default PdfSaveOptions, and saves the file as a monochrome PDF (output_grayscale.pdf).
class Program
{
    static void Main()
    {
        // Create a new workbook and add some sample data
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Sample Data");
        sheet.Cells["A2"].PutValue(123);

        // Apply grayscale (black‑and‑white) printing setting to every worksheet
        foreach (Worksheet ws in workbook.Worksheets)
        {
            ws.PageSetup.BlackAndWhite = true; // forces printing in black and white
        }

        // Configure PDF save options (default options are sufficient for grayscale)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file using the specified options
        workbook.Save("output_grayscale.pdf", pdfOptions);
    }
}
