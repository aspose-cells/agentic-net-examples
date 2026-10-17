// Title: Convert an Excel workbook to a fillable PDF while preserving form fields using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with fillable form fields | How to keep Excel form controls when saving as PDF using Aspose.Cells | C# save .xlsx as PDF preserving interactive form elements | PdfSaveOptions keep form fields in exported PDF Aspose.Cells | Convert workbook to PDF while retaining user input fields .NET
// Tags: Aspose.Cells PDFSaveOptions retain form fields | C# generate fillable PDF from .xlsx using Aspose.Cells | preserve interactive controls when converting Excel to PDF | save workbook as PDF with editable form elements | Aspose.Cells convert workbook to PDF with form preservation

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel workbook, ensures the output folder exists, configures PdfSaveOptions (which keep form fields by default), and saves the workbook as a PDF that retains interactive form fields for user input, with error handling for missing files and other exceptions.
class Program
{
    static void Main()
    {
        // Path to the source Excel workbook
        string excelPath = @"C:\Input\SampleWorkbook.xlsx";

        // Path where the PDF will be saved
        string pdfPath = @"C:\Output\SampleWorkbook.pdf";

        try
        {
            // Verify that the source file exists
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"Input file not found: {excelPath}");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the workbook from the file system
            Workbook workbook = new Workbook(excelPath);

            // Configure PDF save options (form fields are preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: set PDF/A compliance if required
                // Compliance = PdfCompliance.PdfA1b
            };

            // Save the workbook as PDF with the specified options
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("Workbook successfully converted to PDF with interactive form fields.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
