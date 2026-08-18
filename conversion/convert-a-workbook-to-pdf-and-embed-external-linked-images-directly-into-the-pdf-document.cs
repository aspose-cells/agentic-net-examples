// Title: Convert Excel Workbook to PDF with Embedded Images Using Aspose.Cells for .NET (C#)
// Description: Loads an Excel file, validates its path, and saves it as a PDF with PdfSaveOptions.EmbedAttachments enabled, ensuring OLE objects such as embedded pictures are stored inside the PDF. The example also explains that external linked images cannot be embedded with the current Aspose.Cells API and suggests alternative approaches.
// Keywords: Aspose.Cells PDF conversion C# | PdfSaveOptions EmbedAttachments | Excel to PDF with images | embed OLE objects in PDF | .NET workbook to PDF | convert Excel with embedded pictures
// Common Searches: Aspose.Cells embed images when converting Excel to PDF | PdfSaveOptions EmbedAttachments C# example | How to keep pictures inside PDF generated from Excel | Convert .xlsx to PDF with embedded objects Aspose | C# code to save Excel as PDF with attachments
// Developer Intent: Generate a PDF from an Excel workbook and embed any embedded images or OLE objects directly into the PDF using Aspose.Cells for .NET.
// Use Cases: Produce PDF versions of financial reports that retain embedded charts and logos. | Create PDF invoices from Excel templates while ensuring company branding images are included. | Automate batch conversion of multiple .xlsx files to PDFs, preserving all embedded graphics.
// AI Prompts: Write C# code that loads an Excel file, checks its existence, and saves it as a PDF with PdfSaveOptions.EmbedAttachments set to true using Aspose.Cells. | Explain why external linked images cannot be embedded with the current Aspose.Cells API and recommend workarounds or required library updates. | Generate a PowerShell script that scans a folder for .xlsx files and calls a .NET executable to convert each to PDF with embedded attachments.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel file, validates its path, and saves it as a PDF with PdfSaveOptions.EmbedAttachments enabled, ensuring OLE objects such as embedded pictures are stored inside the PDF. The example also explains that external linked images cannot be embedded with the current Aspose.Cells API and suggests alternative approaches.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file (may contain external linked images)
            string sourceFile = "input.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Error: Source file \"{sourceFile}\" not found.");
                return;
            }

            // Load the workbook from the file system
            Workbook workbook = new Workbook(sourceFile);

            // NOTE: The original code attempted to embed linked images using
            // properties (IsLinked, LinkPath, SetImage) that are not available
            // in the current Aspose.Cells version. If embedding of external
            // linked images is required, it must be handled outside of Aspose.Cells
            // or by using a newer library version that supports those members.
            // The essential conversion to PDF works without this optional step.

            // Create PDF save options and enable embedding of attachments.
            // This ensures that any OLE objects (including embedded pictures)
            // are stored directly inside the resulting PDF.
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true
            };

            // Save the workbook as a PDF file with the specified options.
            string pdfFile = "output.pdf";
            workbook.Save(pdfFile, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {pdfFile}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
