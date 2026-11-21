// Title: How to restrict a PDF generated from an Excel workbook to the first 10 pages using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells limit PDF to first 10 pages when saving workbook | Using PdfSaveOptions.PageCount to set maximum page count for PDF conversion in Aspose.Cells | Export Excel to PDF with a 10‑page limit using Aspose.Cells .NET API | Aspose.Cells PDF conversion restrict number of pages in C# example | Set page count ceiling for PDF output from an Excel file with Aspose.Cells
// Tags: PdfSaveOptions.PageCount property | limit PDF pages Aspose.Cells | Excel to PDF page count control C# | Aspose.Cells PDF conversion page limit | C# restrict PDF output page number

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an existing Excel workbook, configures Aspose.Cells PdfSaveOptions with PageCount set to 10 to cap the generated PDF at ten pages, and saves the result as a PDF file, handling missing input files and exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options to limit the output to the first 10 pages
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                PageCount = 10 // Limit PDF to ten pages
            };

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
