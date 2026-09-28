// Title: Save an Aspose.Cells workbook to PDF in C# while preserving leading apostrophes in cell text
// AI Prompts: Generate C# code that creates a workbook, enables QuotePrefixToStyle, writes a value beginning with an apostrophe, and saves the file as PDF using Aspose.Cells. | Insert a verification step that reads the cell's style to confirm the QuotePrefix flag is true before performing the PDF export. | Configure PdfSaveOptions with CheckWorkbookDefaultFont and ExportDocumentStructure to ensure the leading apostrophe renders correctly in the PDF.
// Common Searches: how to keep leading apostrophe visible when exporting Excel to PDF with Aspose.Cells C# | Aspose.Cells QuotePrefixToStyle PDF output example | C# save workbook as PDF preserving literal string apostrophe | PdfSaveOptions CheckWorkbookDefaultFont effect on apostrophe rendering | export document structure Aspose.Cells PDF accessibility
// Tags: Aspose.Cells PDF export QuotePrefix | preserve leading apostrophe Aspose.Cells | PdfSaveOptions CheckWorkbookDefaultFont C# | ExportDocumentStructure Aspose.Cells PDF | C# workbook Save as PDF apostrophe handling

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for PdfSaveOptions

namespace LeadingApostrophePdfDemo
{
    // The example creates a new workbook, enables QuotePrefixToStyle so a leading apostrophe is treated as a style flag, writes "'Aspose" into cell A1, verifies the QuotePrefix flag, configures PdfSaveOptions (CheckWorkbookDefaultFont and ExportDocumentStructure) for proper Unicode rendering, and saves the workbook as LeadingApostropheOutput.pdf, ensuring the leading apostrophe appears correctly in the generated PDF.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Enable QuotePrefixToStyle so that a leading apostrophe is treated as a style flag
            workbook.Settings.QuotePrefixToStyle = true;

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Put a value that starts with a single quote
            // The apostrophe is used in Excel to indicate that the following text is a literal string
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("'Aspose");

            // Verify that the style reflects the QuotePrefix flag
            bool isQuotePrefixSet = cell.GetStyle().QuotePrefix;
            Console.WriteLine($"QuotePrefix flag on cell A1: {isQuotePrefixSet}");

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure the workbook's default font is checked for Unicode characters
                CheckWorkbookDefaultFont = true,

                // Optional: keep document structure for better accessibility
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF
            string pdfPath = "LeadingApostropheOutput.pdf";
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine($"Workbook saved to PDF at '{pdfPath}'.");
            Console.WriteLine("The leading apostrophe should be displayed correctly in the PDF.");
        }
    }
}
