using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Populate some sample data (optional, just for demonstration)
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Fruits");
        sheet.Cells["A3"].PutValue("Vegetables");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["B2"].PutValue(50);
        sheet.Cells["B3"].PutValue(30);

        // Configure PDF conversion options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Embed TrueType fonts in the PDF (required for proper font rendering)
        pdfOptions.EmbedStandardWindowsFonts = true;

        // Retain slicer formatting and document structure in the PDF
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as a PDF using the configured options
        workbook.Save("output.pdf", pdfOptions);
    }
}