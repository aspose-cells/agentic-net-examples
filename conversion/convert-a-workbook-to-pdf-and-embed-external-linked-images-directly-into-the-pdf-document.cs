// Title: Embed Linked Images When Converting Excel to PDF with Aspose.Cells for .NET
// Description: Load an Excel workbook, refresh external linked data, enable PdfSaveOptions.EmbedAttachments, and save the file as a PDF so that all linked images are embedded directly in the PDF output.
// Keywords: Aspose.Cells embed linked images PDF | Excel to PDF conversion .NET | PdfSaveOptions EmbedAttachments | UpdateLinkedDataSource Aspose.Cells | external image linking Excel PDF | C# Aspose.Cells PDF export
// Common Searches: how to embed linked images in PDF using Aspose.Cells | Aspose.Cells update linked data before PDF conversion | PdfSaveOptions EmbedAttachments example C# | convert Excel with external images to PDF .NET | Aspose.Cells PDF/A with embedded graphics
// Developer Intent: Generate a PDF from an Excel workbook that includes all externally linked images as embedded content.
// Use Cases: Create client‑ready PDF reports from Excel templates that contain linked logos or charts. | Batch‑process multiple workbooks, ensuring linked graphics are preserved in archival PDFs. | Produce compliance‑oriented PDFs (e.g., PDF/A) where all media must be self‑contained.
// AI Prompts: Write C# code using Aspose.Cells to load an .xlsx file, refresh external links, and save it as a PDF with embedded linked images. | Explain the effect of PdfSaveOptions.EmbedAttachments on PDF output and how to combine it with PDF/A compliance settings. | Provide a script that scans a directory of Excel files, updates linked data sources, and converts each to a PDF with embedded images using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPdfEmbeddingExample
{
    // Load an Excel workbook, refresh external linked data, enable PdfSaveOptions.EmbedAttachments, and save the file as a PDF so that all linked images are embedded directly in the PDF output.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file that contains external linked images
            string sourcePath = "input.xlsx";

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourcePath);

            // Ensure that any external linked data (including linked images) is refreshed
            // This will pull the latest data from the external sources into the workbook
            // Passing null updates all external links without needing additional source workbooks
            workbook.UpdateLinkedDataSource(null);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // When set to true, Aspose.Cells embeds OLE objects (including linked images) directly into the PDF
            // This ensures that external linked images become part of the PDF file rather than references
            pdfOptions.EmbedAttachments = true;

            // Optional: you can enable other PDF features as needed, e.g., embed fonts, set compliance, etc.
            // pdfOptions.EmbedStandardWindowsFonts = true;
            // pdfOptions.Compliance = PdfCompliance.PdfA1b;

            // Save the workbook as a PDF with the specified options
            string outputPath = "output.pdf";
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF with embedded images: {outputPath}");
        }
    }
}
