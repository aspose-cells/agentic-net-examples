// Title: Setting PDF document properties (Title, Author, Subject) when converting an Excel workbook to PDF with Aspose.Cells in C#
// Common Searches: how to set PDF title and author when saving Excel as PDF using Aspose.Cells C# | Aspose.Cells C# add PDF metadata after workbook conversion | set PDF document properties with Aspose.Cells and Aspose.Pdf in .NET | cannot set PDF metadata via PdfSaveOptions Aspose.Cells workaround | add custom PDF metadata to Excel to PDF conversion in C#
// Tags: Aspose.Cells PDF document properties | C# set PDF metadata after Excel conversion | PdfSaveOptions metadata limitation Aspose.Cells | Aspose.Pdf modify PDF properties post conversion | Excel to PDF metadata handling .NET

using Aspose.Cells;
using System;
using System.IO;

// The example creates a Workbook, writes a value to cell A1, configures PdfSaveOptions, and saves the workbook as a PDF. It highlights that PdfSaveOptions does not provide a way to set PDF metadata such as Title, Author, or Subject directly, and suggests using Aspose.Pdf to add those properties after the conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create

            // Optional: add some data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello World");

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // NOTE: Setting PDF document metadata (Title, Author, Subject) is not supported
            // directly via PdfSaveOptions in the current Aspose.Cells version.
            // If needed, metadata can be added later using Aspose.Pdf after conversion.

            // Save the workbook as PDF
            string outputPath = "SampleOutput.pdf";
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
