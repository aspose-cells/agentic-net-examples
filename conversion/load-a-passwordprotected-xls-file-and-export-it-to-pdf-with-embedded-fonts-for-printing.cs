// Title: C# – Convert a Password‑Protected XLS to PDF with Embedded Fonts using Aspose.Cells
// Description: The example verifies the source file, opens a password‑protected XLS workbook with LoadOptions.Password, configures PdfSaveOptions (fonts are embedded by default), and saves the workbook as a print‑ready PDF while handling possible errors.
// Keywords: Aspose.Cells | C# | password protected XLS | LoadOptions.Password | PdfSaveOptions | embed fonts | Excel to PDF conversion | protected Excel PDF | print ready PDF | Aspose.Cells example
// Common Searches: Aspose.Cells open password protected Excel C# | convert protected xls to pdf with embedded fonts | C# load workbook with password Aspose.Cells | embed fonts when saving Excel as PDF | print ready PDF from Excel using Aspose | PdfSaveOptions font embedding example
// Developer Intent: Open a password‑protected XLS workbook in C# and export it to a PDF that embeds all fonts for reliable printing.
// Use Cases: Automated batch conversion of secured Excel reports to print‑ready PDFs for archival. | Generating PDF invoices from password‑protected Excel templates while preserving exact font appearance. | Providing a web API that accepts protected XLS files and returns PDFs with embedded fonts for downstream printing workflows.
// AI Prompts: Write C# code with Aspose.Cells that opens a password‑protected .xls file and saves it as a PDF with all fonts embedded, including file‑existence checks and exception handling. | Show how to configure PdfSaveOptions in Aspose.Cells to guarantee font embedding and unrestricted printing when converting an Excel workbook to PDF. | Explain the step‑by‑step process for using LoadOptions.Password to load a protected workbook and verify that the resulting PDF contains embedded fonts.

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the source file, opens a password‑protected XLS workbook with LoadOptions.Password, configures PdfSaveOptions (fonts are embedded by default), and saves the workbook as a print‑ready PDF while handling possible errors.
class ConvertProtectedXlsToPdf
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string sourceFile = "protected.xls";

        // Desired output PDF file
        string outputFile = "output.pdf";

        // Password used to protect the XLS workbook
        string workbookPassword = "excelPwd";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file not found: {sourceFile}");
            return;
        }

        try
        {
            // Load the workbook with the password using LoadOptions.Password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = workbookPassword
            };

            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Prepare PDF save options. Embedded fonts are included by default.
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            // By default, the PDF allows printing. No additional security settings are required.
            // If specific security settings are needed, ensure the Aspose.Cells.Pdf assembly is referenced.

            // Save the workbook as PDF with the specified options
            workbook.Save(outputFile, pdfSaveOptions);

            Console.WriteLine("Password‑protected XLS has been converted to PDF with embedded fonts.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
