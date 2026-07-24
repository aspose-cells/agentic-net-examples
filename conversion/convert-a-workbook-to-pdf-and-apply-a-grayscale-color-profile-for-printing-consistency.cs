// Title: C# – Convert Excel to Grayscale PDF with Aspose.Cells
// Description: Load an Excel workbook, enable the BlackAndWhite flag on each worksheet, and save it as a PDF using Aspose.Cells. The output PDF is monochrome, perfect for consistent black‑and‑white printing.
// Keywords: Aspose.Cells | C# Excel to PDF | grayscale PDF | BlackAndWhite | PdfSaveOptions | monochrome export | Excel PDF conversion .NET | print‑ready PDF | page setup black and white | Aspose.Cells PDFSaveOptions
// Common Searches: Aspose.Cells convert Excel to black and white PDF | C# set worksheet BlackAndWhite before PDF export | how to generate grayscale PDF from Excel using Aspose | PdfSaveOptions grayscale output .NET | export Excel as monochrome PDF C#
// Developer Intent: Create a PDF from an Excel file where all pages are rendered in grayscale for reliable printing.
// Use Cases: Generate print‑ready black‑and‑white PDFs for corporate reports. | Automate archival of spreadsheets as grayscale PDFs to reduce storage size. | Provide a web API that returns a monochrome PDF version of an uploaded workbook. | Batch convert multiple workbooks to grayscale PDFs in a scheduled job.
// AI Prompts: Give C# code that converts an .xlsx file to a grayscale PDF using Aspose.Cells, including setting BlackAndWhite for each worksheet. | Write a reusable method ConvertToGrayscalePdf(string inputPath, string outputPath) with Aspose.Cells. | Explain which PdfSaveOptions properties affect color output and how to ensure a PDF is monochrome with Aspose.Cells. | Show how to batch process a folder of Excel files into grayscale PDFs using Aspose.Cells in C#.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Load an Excel workbook, enable the BlackAndWhite flag on each worksheet, and save it as a PDF using Aspose.Cells. The output PDF is monochrome, perfect for consistent black‑and‑white printing.
class ConvertWorkbookToPdfGrayscale
{
    static void Main()
    {
        // Load the source workbook (replace with actual path)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply grayscale (black‑and‑white) printing setting to every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.PageSetup.BlackAndWhite = true;
        }

        // Configure PDF save options (default options are sufficient for grayscale)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file with the grayscale setting applied
        workbook.Save("output.pdf", pdfOptions);
    }
}
