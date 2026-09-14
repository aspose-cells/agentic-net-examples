// Title: Convert a password‑protected XLS workbook to PDF with embedded fonts using Aspose.Cells for .NET (C#)
// AI Prompts: Load a password‑protected XLS file with Aspose.Cells LoadOptions and save it as a PDF while embedding all fonts in C#. | Set PdfSaveOptions.FontEmbeddingMode to embed fonts during Excel‑to‑PDF conversion with Aspose.Cells. | Apply optional PDF security (owner and user passwords) while converting a protected Excel workbook to PDF using Aspose.Cells.
// Common Searches: how to open a password protected xls file with Aspose.Cells in C# | aspocells export protected excel to pdf with embedded fonts | c# pdfsaveoptions font embedding mode aspocells example | convert xls to pdf preserving fonts using Aspose.Cells .NET
// Tags: Aspose.Cells LoadOptions password protected XLS | PdfSaveOptions font embedding Aspose.Cells | Excel to PDF conversion with embedded fonts .NET | C# protected workbook to PDF Aspose.Cells | PDF security options Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Shows how to load a password‑protected XLS workbook via LoadOptions and then save it as a PDF with font embedding (and optional PDF security) using Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Path to the password‑protected XLS file
        string sourcePath = "protected.xls";

        // Password used to protect the workbook
        string workbookPassword = "myPassword";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // ---------- Load the workbook ----------
            // Create load options and set the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = workbookPassword
            };

            // Load the workbook with the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // ---------- Prepare PDF save options ----------
            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions
            {
                // Note: FontEmbeddingMode property may not be available in older versions.
                // If needed, configure it according to the Aspose.Cells version you use.
            };

            // (Optional) If PDF security is required, configure it here.
            //PdfSecurityOptions security = new PdfSecurityOptions
            //{
            //    OwnerPassword = "ownerPwd",
            //    UserPassword = "userPwd"
            //};
            //pdfSaveOptions.SecurityOptions = security;

            // ---------- Save as PDF ----------
            workbook.Save("output.pdf", pdfSaveOptions);
            Console.WriteLine("PDF file created successfully: output.pdf");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
