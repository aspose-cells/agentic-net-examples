using System;
using Aspose.Cells; // Aspose.Cells namespace

class Program
{
    static void Main()
    {
        // Load the source Excel workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Preserve hyperlinks and document structure in the resulting PDF
        pdfOptions.ExportDocumentStructure = true;

        // Save the workbook as PDF using the options (lifecycle rule: Save(string, SaveOptions))
        workbook.Save("output.pdf", pdfOptions);
    }
}