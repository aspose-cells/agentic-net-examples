// Title: Convert an Excel workbook to PDF and render cell comments as footnotes using Aspose.Cells for .NET (comment rendering not supported)
// Common Searches: Aspose.Cells C# export Excel to PDF with cell comments as footnotes | How to include Excel comments in PDF output using Aspose.Cells .NET | C# save workbook as PDF with comments footnote option Aspose | Aspose.Cells PDFSaveOptions comment rendering mode not supported | Convert .xlsx to .pdf preserving comments Aspose.Cells C#
// Tags: Aspose.Cells PDFSaveOptions comment rendering | Excel to PDF conversion with footnote comments C# | preserve cell annotations during .xlsx to .pdf conversion Aspose | comment footnote support limitation Aspose.Cells | C# Aspose.Cells workbook PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example checks for the input XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (noting that comment rendering as footnotes is not supported in the current version), and saves the workbook as a PDF while handling potential exceptions.
class WorkbookToPdfWithComments
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (comment rendering mode not available in this version)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
