// Title: Embed multiple external files as attachments when saving an Excel workbook to PDF with Aspose.Cells for .NET
// AI Prompts: Write C# code that enables PdfSaveOptions.EmbedAttachments and adds a list of file paths so the specified files are embedded as attachments in the PDF produced from an Excel workbook. | Create a helper method that receives an array of attachment file names, configures PdfSaveOptions accordingly, and saves a given Workbook as a PDF containing those attachments. | Show how to combine Aspose.Cells Workbook.Save with PdfSaveOptions to embed both a PDF and a DOCX file as attachments in the output PDF.
// Common Searches: Aspose.Cells C# embed multiple files as attachments in PDF conversion | PdfSaveOptions.EmbedAttachments true how to add attachment paths | Save Excel as PDF with attached documents using Aspose.Cells .NET | C# example for embedding external files into PDF generated from Excel
// Tags: Aspose.Cells PdfSaveOptions embed attachments | C# embed external files in PDF with Aspose.Cells | set EmbedAttachments true Aspose.Cells | add attachment file paths to PdfSaveOptions | Excel to PDF conversion with embedded documents

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, creates a PdfSaveOptions object with EmbedAttachments set to true, adds the desired external file paths to the Attachments collection, and then saves the workbook as a PDF that contains those files as embedded attachments. It also ensures the output folder exists and includes basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string workbookPath = @"C:\Data\Sample.xlsx";

            // Verify workbook file exists
            if (!File.Exists(workbookPath))
                throw new FileNotFoundException("Workbook file not found.", workbookPath);

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Configure PDF save options (default options used)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Ensure output directory exists
            string outputPath = @"C:\Output\Result.pdf";
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
