// Title: Convert XLSB to PDF with UTC Creation Timestamp using Aspose.Cells C#
// Description: Loads an XLSB workbook with Aspose.Cells, sets PdfSaveOptions.CreatedTime to DateTime.UtcNow, and saves the file as a PDF so the document metadata records the current UTC creation time.
// Keywords: Aspose.Cells XLSB to PDF | PdfSaveOptions CreatedTime | set PDF creation date C# | export XLSB as PDF with metadata | UTC timestamp PDF Aspose.Cells | C# Excel binary to PDF conversion
// Common Searches: Aspose.Cells set PDF creation time | C# convert XLSB to PDF with timestamp | PdfSaveOptions CreatedTime example | how to add UTC metadata to PDF using Aspose.Cells | batch XLSB to PDF conversion with creation date
// Developer Intent: Generate a PDF from an XLSB workbook and embed the current UTC time as the PDF's creation metadata.
// Use Cases: Produce audit‑ready PDFs from XLSB reports with precise creation timestamps for compliance. | Automate batch conversion of XLSB files to PDFs while recording the processing time in the document metadata. | Create time‑stamped PDF invoices or statements from XLSB templates in a .NET backend service.
// AI Prompts: Write C# code with Aspose.Cells that converts an XLSB file to PDF and sets the PDF's CreatedTime to DateTime.UtcNow. | Show how to configure PdfSaveOptions to add custom metadata such as author, title, and UTC creation date when saving a workbook as PDF. | Explain a method to batch process multiple XLSB files into PDFs, ensuring each PDF includes the conversion timestamp in its metadata.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // For PdfSaveOptions

// Loads an XLSB workbook with Aspose.Cells, sets PdfSaveOptions.CreatedTime to DateTime.UtcNow, and saves the file as a PDF so the document metadata records the current UTC creation time.
class XlsbToPdfConverter
{
    static void Main()
    {
        // Path to the source XLSB file
        string sourcePath = "input.xlsb";

        // Path for the resulting PDF file
        string destPath = "output.pdf";

        // Load the XLSB workbook
        Workbook workbook = new Workbook(sourcePath);

        // Configure PDF save options and set the creation time to current UTC
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            CreatedTime = DateTime.UtcNow
        };

        // Save the workbook as PDF using the specified options
        workbook.Save(destPath, pdfOptions);
    }
}
