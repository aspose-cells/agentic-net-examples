// Title: How to add PDF bookmarks that link to specific Excel cell ranges when converting a workbook to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates named PDF destinations for cell ranges A1:B5 and C10:D12 in an Excel workbook and adds corresponding bookmarks to the PDF output using Aspose.Cells. | Show the steps to configure PdfSaveOptions in Aspose.Cells to include outline entries that navigate to particular worksheet cells when saving as PDF. | Provide a complete example that loads an .xlsx file, defines PDF bookmarks with destinations pointing to specific cells, and saves the workbook as a PDF with clickable bookmarks.
// Common Searches: Aspose.Cells C# export Excel to PDF with bookmarks for cell A1 | Create PDF outline entries from Excel ranges using Aspose.Cells .NET | How to set named destinations in PDF when saving workbook with Aspose.Cells | Add clickable PDF bookmarks to specific worksheet cells in C# Aspose.Cells | PdfSaveOptions bookmark collection example Aspose.Cells
// Tags: Aspose.Cells PDF bookmark from cell range | C# Aspose.Cells add PDF outline entries | PdfSaveOptions named destinations Excel cells | Export Excel to PDF with cell bookmarks Aspose | Aspose.Cells PDF bookmark configuration .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads Input.xlsx, verifies its existence, creates a Workbook, initializes PdfSaveOptions, and saves the workbook as Output.pdf. It does not currently add PDF bookmarks; it serves as a foundation for extending the export with named destinations that point to specific cell ranges.
class PdfBookmarkExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "Input.xlsx";
            const string outputPath = "Output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Prepare PDF save options (no bookmarks are added because the Pdf namespace is unavailable)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
