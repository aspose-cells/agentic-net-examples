using System;
using Aspose.Cells;
using Aspose.Cells.Saving; // PdfSaveOptions resides here

class HtmlToPdfConverter
{
    static void Main()
    {
        // Path to the source HTML file
        string htmlFilePath = "input.html";

        // Load the HTML document into a Workbook using HtmlLoadOptions
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        Workbook workbook = new Workbook(htmlFilePath, loadOptions);

        // Configure PDF save options (optional settings can be adjusted here)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            // Preserve the document structure for better accessibility
            ExportDocumentStructure = true
        };

        // Save the workbook directly as a PDF file, preserving layout
        string pdfOutputPath = "output.pdf";
        workbook.Save(pdfOutputPath, pdfOptions);

        Console.WriteLine($"HTML file '{htmlFilePath}' has been converted to PDF '{pdfOutputPath}'.");
    }
}