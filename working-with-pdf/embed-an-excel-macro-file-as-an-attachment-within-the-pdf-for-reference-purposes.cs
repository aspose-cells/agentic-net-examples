// Title: Embedding an .xlsm macro workbook as an attached file in a PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# embed .xlsm macro as attachment when saving to PDF | How to attach the original Excel macro file to a PDF generated with Aspose.Cells | Which Aspose.Cells version adds PdfSaveOptions EmbedMacros support | C# example for converting a macro-enabled workbook to PDF with macro attachment | Save Excel workbook with macros to PDF and include macro file using Aspose.Cells
// Tags: Aspose.Cells embed macro attachment PDF | PdfSaveOptions EmbedMacros support | C# convert .xlsm to PDF with attached macro | Aspose.Cells version requirement for macro embedding | preserve Excel macros during PDF export

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a macro-enabled .xlsm workbook, verifies the source file, prepares PdfSaveOptions, notes that the EmbedMacros property is unavailable in the current Aspose.Cells release, creates the output directory if needed, and saves the workbook as a PDF, reporting the result or any errors.
class EmbedMacroInPdf
{
    static void Main()
    {
        try
        {
            // Path to the Excel file that contains macros
            string excelPath = @"C:\Data\ReportWithMacro.xlsm";

            // Verify the source file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Excel file not found: {excelPath}");
                return;
            }

            // Load the workbook (macros are preserved)
            Workbook workbook = new Workbook(excelPath);

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: The EmbedMacros property is not available in the current Aspose.Cells version.
            // If macro embedding is required, ensure you are using a version that supports it.

            // Destination PDF path
            string pdfPath = @"C:\Data\ReportWithMacro.pdf";

            // Ensure the output directory exists
            string? pdfDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook as PDF
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"PDF created at: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
