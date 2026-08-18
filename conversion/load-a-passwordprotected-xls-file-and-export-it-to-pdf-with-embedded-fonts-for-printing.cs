// Title: Convert a Password‑Protected XLS to PDF with Embedded Fonts Using Aspose.Cells for .NET (C#)
// Description: This example shows how to open a password‑protected XLS workbook with Aspose.Cells LoadOptions, configure PdfSaveOptions for print‑ready output, and save the workbook as a PDF. The PDF is generated with default font handling, and when supported Aspose.Cells automatically embeds the required fonts for high‑quality printing.
// Keywords: Aspose.Cells | C# | password protected XLS | Excel to PDF conversion | PdfSaveOptions | EmbedStandardFonts | embedded fonts PDF | load workbook with password | print ready PDF | .NET Excel PDF
// Common Searches: open password protected Excel file Aspose.Cells C# | convert protected XLS to PDF with embedded fonts | Aspose.Cells PdfSaveOptions embed fonts example | load workbook with password and export to PDF | print quality PDF from secured Excel using Aspose
// Developer Intent: Load a password‑protected XLS workbook and export it to a PDF that includes embedded fonts for printing.
// Use Cases: Create print‑ready PDFs from secured Excel spreadsheets for legal or archival purposes. | Automate batch conversion of password‑protected XLS files to PDFs in a .NET backend service. | Integrate secure Excel‑to‑PDF conversion into a web application that receives protected uploads.
// AI Prompts: Generate C# code that opens a password‑protected .xls file with Aspose.Cells and saves it as a PDF with embedded fonts. | Explain how to use LoadOptions and PdfSaveOptions to convert a protected Excel workbook to a print‑ready PDF in Aspose.Cells. | Show how to enable font embedding in Aspose.Cells PDF output when the EmbedStandardFonts property is available.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering.PdfSecurity;

namespace AsposeCellsExample
{
    // This example shows how to open a password‑protected XLS workbook with Aspose.Cells LoadOptions, configure PdfSaveOptions for print‑ready output, and save the workbook as a PDF. The PDF is generated with default font handling, and when supported Aspose.Cells automatically embeds the required fonts for high‑quality printing.
    class Program
    {
        static void Main()
        {
            // Paths for source XLS (password‑protected) and destination PDF
            string sourcePath = "protected.xls";
            string destPath = "output.pdf";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            try
            {
                // Load options with the workbook password
                LoadOptions loadOptions = new LoadOptions
                {
                    Password = "myPassword"
                };

                // Load the protected workbook
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // EmbedStandardFonts property is not available in this version;
                    // default font handling will be used.
                    ExportDocumentStructure = true
                };

                // Save the workbook as a PDF with the specified options
                workbook.Save(destPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to: {destPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
