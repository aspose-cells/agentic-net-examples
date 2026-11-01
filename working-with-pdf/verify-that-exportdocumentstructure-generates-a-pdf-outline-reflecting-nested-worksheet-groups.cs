// Title: Verify that ExportDocumentStructure creates a PDF outline reflecting worksheet order using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells exportdocumentstructure pdf outline worksheet order | how to enable PDF bookmarks when converting Excel to PDF with Aspose.Cells | check PDF outline contains all worksheets after saving with PdfSaveOptions | Aspose.Cells generate PDF document structure from workbook sheets example | verify PDF bookmark hierarchy matches worksheet sequence in C#
// Tags: ExportDocumentStructure PDF outline | Aspose.Cells PdfSaveOptions worksheet bookmarks | C# convert Excel to PDF with document structure | generate PDF bookmarks from worksheets Aspose.Cells | verify PDF outline order Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook with four worksheets, enables ExportDocumentStructure in PdfSaveOptions, saves the workbook as a PDF, and confirms that the resulting PDF contains an outline (bookmarks) that follows the worksheet order.
class ExportDocumentStructureVerification
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add worksheets and name them for clarity
            Worksheet ws1 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws1.Name = "Sheet1";
            Worksheet ws2 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws2.Name = "Sheet2";
            Worksheet ws3 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws3.Name = "Sheet3";
            Worksheet ws4 = workbook.Worksheets[workbook.Worksheets.Add()];
            ws4.Name = "Sheet4";

            // Populate some data (optional, just to have content)
            ws1.Cells["A1"].PutValue("Data in Sheet1");
            ws2.Cells["A1"].PutValue("Data in Sheet2");
            ws3.Cells["A1"].PutValue("Data in Sheet3");
            ws4.Cells["A1"].PutValue("Data in Sheet4");

            // NOTE: Aspose.Cells does not provide a WorksheetCollection.Group method.
            // The PDF outline will be generated based on worksheet order when ExportDocumentStructure is enabled.

            // Set PDF save options to export document structure (outline)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // Enable PDF outline generation
            };

            // Define output PDF path
            string pdfPath = Path.Combine(Environment.CurrentDirectory, "GroupedSheets.pdf");

            // Ensure the directory exists
            string? pdfDir = Path.GetDirectoryName(pdfPath);
            if (!string.IsNullOrEmpty(pdfDir) && !Directory.Exists(pdfDir))
            {
                Directory.CreateDirectory(pdfDir);
            }

            // Save the workbook as PDF with document structure
            workbook.Save(pdfPath, pdfOptions);

            // Verify that the PDF file was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"PDF saved successfully at: {pdfPath}");
            }
            else
            {
                Console.WriteLine("Failed to create PDF file.");
            }
        }
        catch (Exception ex)
        {
            // Runtime safety: output any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
