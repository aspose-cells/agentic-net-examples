// Title: Convert an Excel workbook to PDF with each worksheet starting on a new page using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file and saves it as a PDF where each worksheet begins on its own page by enabling single‑page‑per‑sheet pagination in Aspose.Cells. | Show how to configure Aspose.Cells PdfSaveOptions to enforce one‑page‑per‑sheet pagination when converting Excel to PDF in a .NET application. | Create a script that processes a folder of Excel files and exports each to a PDF with separate pages per worksheet using Aspose.Cells.
// Common Searches: Aspose.Cells C# export each Excel sheet to a separate PDF page | PdfSaveOptions enable single page per worksheet example | Convert multiple worksheets to individual pages in PDF with Aspose.Cells | C# code to save workbook as PDF with sheet‑level pagination
// Tags: Aspose.Cells PDF export per worksheet page | C# set PdfSaveOptions for sheet pagination | Excel workbook to PDF with separate pages per sheet | Aspose.Cells configure PDF page breaks per sheet | One-page-per-sheet PDF conversion using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads 'input.xlsx', sets PdfSaveOptions.OnePagePerSheet to true so each worksheet starts on a new PDF page, and saves the workbook as 'output.pdf' using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load an existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to start each worksheet on a new page
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = true // Ensures each sheet begins on a separate PDF page
        };

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
