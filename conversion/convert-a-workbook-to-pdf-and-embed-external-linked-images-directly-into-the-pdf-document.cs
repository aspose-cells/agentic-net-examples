using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Create PDF save options
        PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

        // Enable embedding of external linked images (treated as attachments)
        pdfSaveOptions.EmbedAttachments = true;

        // Save the workbook as PDF with the specified options
        workbook.Save("output.pdf", pdfSaveOptions);
    }
}