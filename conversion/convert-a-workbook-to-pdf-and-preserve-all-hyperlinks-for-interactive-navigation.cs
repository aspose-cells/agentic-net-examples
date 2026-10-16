// Title: Convert an Excel workbook to PDF with clickable hyperlinks using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert .xlsx to PDF preserving hyperlink navigation | save Excel workbook as PDF with active links using PdfSaveOptions in .NET | C# code to keep Excel cell hyperlinks when exporting to PDF with Aspose | how to retain clickable URLs in PDF generated from Excel using Aspose.Cells
// Tags: Aspose.Cells PDF conversion with hyperlink preservation | PdfSaveOptions retain hyperlinks | Excel to PDF clickable links C# | Aspose.Cells workbook to PDF preserving navigation | C# export Excel hyperlinks to PDF

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLSX file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions (which retain hyperlinks by default), and saves the workbook as a PDF, ensuring that all hyperlinks remain interactive in the resulting document.
class WorkbookToPdfConverter
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options preserve hyperlinks)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during conversion: {ex.Message}");
        }
    }
}
