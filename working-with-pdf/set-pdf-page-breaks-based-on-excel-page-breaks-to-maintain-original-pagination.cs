// Title: Convert Excel workbook to PDF while preserving original page breaks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF keep original pagination | how to retain Excel page breaks when saving as PDF with Aspose.Cells | PdfSaveOptions OnePagePerSheet false example in C# | convert Excel sheet to multi‑page PDF based on page breaks Aspose | preserve Excel pagination during PDF conversion using Aspose.Cells .NET
// Tags: Aspose.Cells PDF export page break handling | PdfSaveOptions OnePagePerSheet configuration | C# Excel to PDF multi-page conversion | preserve workbook pagination in PDF | export Excel sheet with original page breaks

using Aspose.Cells;
using System;

// The example loads an Excel workbook, sets PdfSaveOptions.OnePagePerSheet to false so that each existing Excel page break becomes a separate PDF page, and saves the workbook as a PDF, preserving the original pagination.
class Program
{
    static void Main()
    {
        // Load the Excel workbook that contains the original page breaks
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to preserve the original pagination.
        // Setting OnePagePerSheet to false ensures that each Excel page break
        // becomes a separate PDF page rather than forcing the entire sheet onto one page.
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            OnePagePerSheet = false
        };

        // Save the workbook as PDF while maintaining the Excel page breaks.
        workbook.Save("output.pdf", pdfOptions);
    }
}
