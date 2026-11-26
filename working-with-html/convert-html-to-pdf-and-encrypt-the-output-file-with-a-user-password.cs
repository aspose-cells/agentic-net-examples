// Title: Convert an HTML file to an encrypted PDF with a user password using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells convert HTML to PDF and set a user password | How to apply PdfSecurityOptions when saving a PDF from an HTML workbook in Aspose.Cells | Encrypt PDF generated from HTML using Aspose.Cells C# example | Command‑line HTML to password‑protected PDF conversion with Aspose.Cells | Set print and modify permissions on PDF saved from HTML with Aspose.Cells
// Tags: Aspose.Cells HTML to PDF conversion with encryption | PdfSecurityOptions user password C# | PdfSaveOptions set PDF permissions Aspose.Cells | LoadOptions HTML format workbook Aspose.Cells | command‑line HTML to PDF Aspose.Cells sample

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace HtmlToPdfWithEncryption
{
    // This C# example loads an HTML file into an Aspose.Cells Workbook, configures PdfSaveOptions with PdfSecurityOptions to apply a user password and specific permissions, and saves the workbook as an encrypted PDF. Optional command‑line arguments let you specify input HTML, output PDF, and the password.
    class Program
    {
        static int Main(string[] args)
        {
            // Define default file names; they can be overridden via command‑line arguments.
            string inputHtmlPath = "input.html";
            string outputPdfPath = "output.pdf";
            string userPassword = "UserPass123";

            // Simple command‑line parsing: 
            //   args[0] = input HTML file (optional)
            //   args[1] = output PDF file (optional)
            //   args[2] = user password (optional)
            try
            {
                if (args.Length > 0) inputHtmlPath = args[0];
                if (args.Length > 1) outputPdfPath = args[1];
                if (args.Length > 2) userPassword = args[2];
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error parsing arguments: {ex.Message}");
                return 1;
            }

            // Validate input file existence.
            if (!File.Exists(inputHtmlPath))
            {
                Console.WriteLine($"Input HTML file not found: {inputHtmlPath}");
                return 1;
            }

            try
            {
                // Load the HTML file into a workbook.
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Html);
                Workbook workbook = new Workbook(inputHtmlPath, loadOptions);

                // Prepare PDF save options.
                PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

                // Configure PDF security (user password only; owner password left empty).
                PdfSecurityOptions securityOptions = new PdfSecurityOptions
                {
                    UserPassword = userPassword,
                    // OwnerPassword can be set if you want full‑access password.
                    // OwnerPassword = "OwnerSecret",
                    // Example permissions – adjust as needed.
                    PrintPermission = true,
                    ModifyDocumentPermission = false,
                    ExtractContentPermission = false,
                    AnnotationsPermission = true,
                    FillFormsPermission = true,
                    FullQualityPrintPermission = true
                };

                pdfSaveOptions.SecurityOptions = securityOptions;

                // Save the workbook as an encrypted PDF.
                workbook.Save(outputPdfPath, pdfSaveOptions);

                Console.WriteLine($"HTML file \"{inputHtmlPath}\" successfully converted to encrypted PDF \"{outputPdfPath}\".");
                Console.WriteLine($"User password required to open the PDF: \"{userPassword}\"");
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during conversion or encryption: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
                return 1;
            }
        }
    }
}
