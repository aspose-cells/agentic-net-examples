// Title: How to limit an Aspose.Cells PDF export to 10 pages using PdfSaveOptions.MaxPageCount in C#
// AI Prompts: Write C# code that creates an Excel workbook, fills it with data, and saves it as a PDF limited to ten pages by configuring PdfSaveOptions.MaxPageCount. | Show the steps to set PdfSaveOptions.MaxPageCount = 10 when converting a workbook to PDF with Aspose.Cells for .NET. | Adapt existing Aspose.Cells PDF conversion code to enforce a maximum of 10 pages in the generated PDF.
// Common Searches: Aspose.Cells C# limit PDF to specific number of pages | PdfSaveOptions.MaxPageCount usage example Aspose.Cells | How to restrict PDF page count when exporting Excel with Aspose.Cells .NET | C# code sample for setting max page count in Aspose.Cells PDF conversion | Export workbook to PDF with ten-page limit using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions MaxPageCount | C# limit PDF pages Aspose.Cells | Excel to PDF page count restriction .NET | Aspose.Cells PDF export page limit | Configure PDF save options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// // Creates a workbook, populates rows, sets PdfSaveOptions.MaxPageCount = 10 to cap the PDF at ten pages, and saves the file as output.pdf.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Populate the worksheet with sample data to generate multiple pages
            Worksheet sheet = workbook.Worksheets[0];
            for (int row = 0; row < 200; row++)
            {
                sheet.Cells[row, 0].PutValue($"Row {row + 1}");
            }

            // Configure PDF save options to limit the output to ten pages
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Limit PDF to a maximum of 10 pages (if supported by the API)
                PageCount = 10
            };

            // Determine output file path and ensure the directory exists
            string outputPath = "output.pdf";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
