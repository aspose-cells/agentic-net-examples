using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

class WorkbookToPdfWithoutExternalImages
{
    static void Main()
    {
        // Path to the source Excel file (can contain external linked images)
        string sourcePath = "input.xlsx";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        // Load the workbook from the source file
        Workbook workbook = new Workbook(sourcePath);

        // Configure PDF save options
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure that OLE attachments are NOT embedded (default is false,
        // but set explicitly for clarity). External linked images are not
        // embedded in the PDF when this option is false.
        pdfOptions.EmbedAttachments = false;

        // Optional: ignore blank pages to reduce size further
        pdfOptions.PrintingPageType = PrintingPageType.IgnoreBlank;

        // Save the workbook as PDF using the configured options
        workbook.Save(pdfPath, pdfOptions);

        Console.WriteLine($"Workbook converted to PDF without embedding external images: {pdfPath}");
    }
}