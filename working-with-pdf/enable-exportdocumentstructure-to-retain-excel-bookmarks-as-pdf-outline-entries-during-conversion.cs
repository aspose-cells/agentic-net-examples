// Title: How to export Excel workbook bookmarks as PDF outline entries using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, sets PdfSaveOptions.ExportDocumentStructure to true, and saves it as a PDF while preserving the workbook's bookmarks as outline entries. | Describe the configuration steps for Aspose.Cells PdfSaveOptions to retain Excel bookmarks when converting a workbook to PDF in a .NET application.
// Common Searches: Aspose.Cells C# export Excel bookmarks to PDF outline | PdfSaveOptions ExportDocumentStructure true example | Preserve workbook bookmarks during Excel to PDF conversion using Aspose | How to keep Excel bookmark hierarchy in generated PDF with Aspose.Cells | C# code sample for saving .xlsx as PDF with document structure enabled
// Tags: Aspose.Cells ExportDocumentStructure flag | retain workbook bookmarks in PDF | C# PdfSaveOptions bookmark export | Excel to PDF outline generation | document structure option Aspose.Cells

using Aspose.Cells;

// The example loads an Excel workbook, enables the ExportDocumentStructure flag in PdfSaveOptions to export workbook bookmarks as PDF outline entries, and saves the result as a PDF file.
class Program
{
    static void Main()
    {
        // Load the Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure PDF save options to export document structure (bookmarks become PDF outline entries)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // Enable bookmark export

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}
