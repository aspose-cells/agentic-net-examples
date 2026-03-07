using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create HTML load options (using the provided constructor rule)
        HtmlLoadOptions loadOptions = new HtmlLoadOptions();
        // Example: enable support for <div> tags while loading
        loadOptions.SupportDivTag = true;

        // Load the HTML file into a workbook (using the provided load rule)
        Workbook workbook = new Workbook("input.html", loadOptions);

        // Create PDF save options (property ExportDocumentStructure is covered by a rule)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        pdfOptions.ExportDocumentStructure = true; // retain document structure in the PDF

        // Save the workbook as a PDF file (using the provided save rule)
        workbook.Save("output.pdf", pdfOptions);
    }
}