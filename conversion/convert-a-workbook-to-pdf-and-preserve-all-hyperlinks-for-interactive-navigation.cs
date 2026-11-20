// Title: Convert an Excel workbook to PDF with clickable hyperlinks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF preserving hyperlink clicks | How to keep Excel hyperlinks active when saving as PDF with Aspose | C# PdfSaveOptions keep links in PDF generated from .xlsx | Generate PDF from workbook that retains hyperlink navigation using Aspose.Cells | Save Excel file as PDF with clickable URLs in .NET
// Tags: Aspose.Cells PDFSaveOptions hyperlink support | C# Excel workbook to PDF with active links | preserve clickable URLs during PDF conversion | Aspose.Cells generate PDF with hyperlink navigation | hyperlink-enabled PDF export using .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an existing XLSX file with Aspose.Cells, creates a PdfSaveOptions object (hyperlinks are retained by default), and saves the workbook as a PDF, ensuring all embedded hyperlinks remain functional for interactive navigation.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (hyperlinks are preserved by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
