// Title: How to embed PDF title, author, and subject metadata when saving an Aspose.Cells workbook to PDF in C#
// AI Prompts: Generate a PDF from an Aspose.Cells workbook and embed custom Title, Author, and Subject properties using C#. | Demonstrate setting built‑in document properties before calling Workbook.Save with PdfSaveOptions in Aspose.Cells. | Add additional PDF metadata such as Keywords and Creator to a workbook‑to‑PDF conversion with Aspose.Cells.
// Common Searches: Aspose.Cells C# set PDF document title author subject before saving | How to add custom metadata to PDF generated from Excel using Aspose.Cells | Embedding PDF metadata when converting workbook to PDF with PdfSaveOptions in .NET | C# code example for setting built‑in document properties for PDF output in Aspose.Cells
// Tags: Aspose.Cells set PDF metadata | PdfSaveOptions embed document properties | C# workbook to PDF with custom title | built‑in document properties Aspose.Cells | Excel to PDF conversion metadata .NET

using Aspose.Cells;
using System;
using System.IO;

// // Creates a workbook, assigns Title, Author, and Subject via BuiltInDocumentProperties, and saves it as a PDF using PdfSaveOptions, embedding the specified metadata.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data (optional)
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample content for PDF");

            // Set document properties (metadata) that will be written to the PDF
            workbook.BuiltInDocumentProperties.Title = "My PDF Title";
            workbook.BuiltInDocumentProperties.Author = "Jane Smith";
            workbook.BuiltInDocumentProperties.Subject = "PDF metadata example";

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Define output file path
            string outputPath = "MyWorkbook.pdf";

            // Save the workbook as PDF with the specified metadata
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
