// Title: Export an Aspose.Cells workbook to PDF with cell comments as PDF annotations (C#)
// Description: Demonstrates how to create a workbook, add data and a comment, enable ExportDocumentStructure in PdfSaveOptions, and save the file as a PDF where the comment appears as an annotation.
// Keywords: Aspose.Cells PDF export | C# export comments to PDF | ExportDocumentStructure | cell comments as PDF annotations | Excel to PDF with notes | Aspose.Cells PdfSaveOptions
// Common Searches: Aspose.Cells keep Excel comments when converting to PDF | PdfSaveOptions ExportDocumentStructure example C# | save workbook as PDF with annotations Aspose | export cell notes to PDF using Aspose.Cells .NET | convert Excel to PDF preserving comments
// Developer Intent: Generate a PDF from an Excel workbook while retaining each cell comment as a searchable PDF annotation.
// Use Cases: Produce financial statements where analyst remarks stay attached to cells in the PDF. | Create product catalogs that include tax or disclaimer notes stored as comments. | Automate bulk conversion of spreadsheets to PDFs without losing reviewer feedback.
// AI Prompts: Provide a C# snippet that saves an Aspose.Cells workbook to PDF with comments exported as annotations. | Show how to add multiple cell comments and verify they appear as PDF annotations after using ExportDocumentStructure. | Explain the role of PdfSaveOptions.ExportDocumentStructure and how to test that comments are preserved in the output PDF.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfWithComments
{
    // Demonstrates how to create a workbook, add data and a comment, enable ExportDocumentStructure in PdfSaveOptions, and save the file as a PDF where the comment appears as an annotation.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some sample data
                sheet.Cells["A1"].PutValue("Product");
                sheet.Cells["B1"].PutValue("Price");
                sheet.Cells["A2"].PutValue("Apple");
                sheet.Cells["B2"].PutValue(1.20);
                sheet.Cells["A3"].PutValue("Banana");
                sheet.Cells["B3"].PutValue(0.80);

                // Add a comment to a cell (will be exported as a PDF annotation)
                // Aspose.Cells requires adding the comment first, then setting its properties
                int commentIndex = sheet.Comments.Add("B2");
                Comment comment = sheet.Comments[commentIndex];
                comment.Author = "John Doe";
                comment.Note = "Price includes tax.";

                // Create PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Enable document structure export – this preserves comments as PDF annotations
                    ExportDocumentStructure = true
                };

                // Optional: calculate formulas before saving (not needed here but good practice)
                workbook.CalculateFormula();

                // Save the workbook as PDF with the specified options
                string outputPath = "WorkbookWithComments.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF file created at '{outputPath}' with cell comments preserved as annotations.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
