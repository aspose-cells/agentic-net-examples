// Title: Convert Aspose.Cells Workbook to PDF with System Font Substitution (C#)
// Description: Shows how to enable system‑font fallback in Aspose.Cells by setting FontConfigs.PreferSystemFontSubstitutes and PdfSaveOptions.CheckWorkbookDefaultFont, then saving a workbook with Unicode text to PDF.
// Keywords: Aspose.Cells PDF conversion C# | FontConfigs.PreferSystemFontSubstitutes | PdfSaveOptions.CheckWorkbookDefaultFont | system font fallback | missing fonts Aspose.Cells | Unicode Excel to PDF | server‑side Excel PDF export | C# Aspose.Cells example
// Common Searches: Aspose.Cells enable system font substitution | PdfSaveOptions CheckWorkbookDefaultFont example C# | C# convert Excel to PDF when fonts are missing | fallback to system font Aspose.Cells PDF export | Unicode characters PDF conversion Aspose.Cells
// Developer Intent: Generate a PDF from an Excel workbook that automatically uses the operating system’s default font if the original font is unavailable.
// Use Cases: Export multilingual reports on servers that only have system fonts installed. | Create PDF documents in CI/CD pipelines where custom fonts cannot be guaranteed. | Batch‑process legacy Excel files that reference fonts not present on the target machine.
// AI Prompts: Write C# code with Aspose.Cells to save a workbook as PDF using system font substitution for missing fonts. | Explain the interaction between FontConfigs.PreferSystemFontSubstitutes and PdfSaveOptions.CheckWorkbookDefaultFont in Aspose.Cells. | Provide step‑by‑step instructions to convert an Excel file containing Unicode text to PDF, ensuring the system default font is used when required fonts are absent.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // Shows how to enable system‑font fallback in Aspose.Cells by setting FontConfigs.PreferSystemFontSubstitutes and PdfSaveOptions.CheckWorkbookDefaultFont, then saving a workbook with Unicode text to PDF.
    public class WorkbookToPdfWithSystemFontSubstitution
    {
        public static void Run()
        {
            try
            {
                // Prefer system font substitutes when the original font is missing
                FontConfigs.PreferSystemFontSubstitutes = true;

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add sample data that may require font substitution
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample text with Unicode: 你好, مرحبا, Привет");

                // Configure PDF save options to use the workbook's default font checking
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // When true, Aspose.Cells will try to use the workbook's default font first.
                    // If the font is unavailable, it will fall back to the system default font.
                    CheckWorkbookDefaultFont = true
                    // DefaultFont left unset to use system default.
                };

                // Save the workbook as PDF using the configured options
                workbook.Save("output.pdf", pdfOptions);

                Console.WriteLine("Workbook has been saved to PDF with system default font substitution.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            WorkbookToPdfWithSystemFontSubstitution.Run();
        }
    }
}
