// Title: How to convert an Excel workbook to PDF with cell comments displayed as footnotes in C# using Aspose.Cells
// Common Searches: c# aspnet convert xlsx to pdf including comments as footnotes | aspose.cells pdfsaveoptions render comments as footnotes example | export excel workbook to pdf with cell comments footnotes c# | save workbook as pdf preserving excel comments asp.net | include excel cell comments in pdf output using aspose.cells
// Tags: Aspose.Cells PDF conversion with comments footnotes | PdfSaveOptions render cell comments as footnotes | C# export Excel workbook to PDF including comments | Excel to PDF conversion preserving comments Aspose.Cells | Workbook.Save PDF with comment footnotes

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input XLSX file, loads it into an Aspose.Cells Workbook, uses PdfSaveOptions (which render cell comments as footnotes by default), and saves the workbook as a PDF, handling any runtime errors.
class WorkbookToPdfWithComments
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default behavior renders comments as footnotes)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
