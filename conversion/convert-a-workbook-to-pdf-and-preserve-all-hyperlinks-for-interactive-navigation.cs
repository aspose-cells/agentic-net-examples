// Title: Aspose.Cells C# – Convert Excel to PDF while preserving clickable hyperlinks
// Description: Demonstrates loading or creating an Excel workbook, adding a hyperlink, setting PdfSaveOptions.ExportDocumentStructure to true, and saving the file as a PDF so that all hyperlinks remain active for interactive navigation.
// Keywords: Aspose.Cells PDF hyperlink preservation | C# Excel to PDF with active links | ExportDocumentStructure true | save Excel as PDF clickable URLs | Aspose.Cells PdfSaveOptions example
// Common Searches: keep hyperlinks when converting Excel to PDF Aspose.Cells | C# Aspose.Cells export PDF with clickable links | how to preserve Excel hyperlinks in PDF | Aspose.Cells PDF hyperlink support
// Developer Intent: Create a PDF from an Excel workbook that retains every cell hyperlink for end‑user navigation.
// Use Cases: Generate PDF reports from spreadsheets where users can click links to external resources. | Automate production of PDF manuals from Excel templates that contain embedded URLs. | Batch‑process multiple workbooks into PDFs while maintaining hyperlink functionality.
// AI Prompts: Write C# code using Aspose.Cells to convert an Excel file to PDF and keep all hyperlinks clickable. | Explain how PdfSaveOptions.ExportDocumentStructure influences hyperlink preservation in the resulting PDF. | Show a sample that adds several hyperlinks to different cells before exporting the workbook to PDF.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfHyperlinkDemo
{
    // Demonstrates loading or creating an Excel workbook, adding a hyperlink, setting PdfSaveOptions.ExportDocumentStructure to true, and saving the file as a PDF so that all hyperlinks remain active for interactive navigation.
    class Program
    {
        static void Main()
        {
            try
            {
                // Path to the source Excel file
                string sourcePath = "input.xlsx";

                // Path to the output PDF file
                string pdfPath = "output.pdf";

                // Ensure the source file exists; create a simple workbook if it does not.
                if (!File.Exists(sourcePath))
                {
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                    tempWb.Save(sourcePath);
                }

                // Load the workbook
                Workbook workbook = new Workbook(sourcePath);

                // Add a sample hyperlink to demonstrate preservation.
                Worksheet sheet = workbook.Worksheets[0];
                // Set the display text in the cell.
                sheet.Cells["A1"].PutValue("Example Site");
                // Add hyperlink to the cell range (A1).
                sheet.Hyperlinks.Add(0, 0, 1, 1, "https://www.example.com");

                // Configure PDF save options to keep hyperlinks active.
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    ExportDocumentStructure = true
                };

                // Save the workbook as PDF
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"Workbook converted to PDF with hyperlinks preserved: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
