// Title: How to fit all worksheet columns onto a single PDF page when converting an Excel file to PDF with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that sets PageSetup.FitToPagesWide = 1 to force column scaling during PDF export with Aspose.Cells. | Demonstrate assigning the worksheet's print area to the maximum display range prior to PDF conversion using Aspose.Cells. | Create a complete C# example that loads an .xlsx, applies column fitting, defines the print area, and saves to PDF with PdfSaveOptions.
// Common Searches: Aspose.Cells C# limit Excel column width to a single PDF page during conversion | how to set FitToPagesWide when exporting Excel to PDF using Aspose.Cells | C# Aspose.Cells PDF export keep rows on multiple pages while columns fit one page | set print area to used range Aspose.Cells PDF output C#
// Tags: PageSetup.FitToPagesWide column scaling Aspose.Cells | define print area used range Aspose.Cells | PdfSaveOptions OnePagePerSheet false .NET | Excel to PDF column fit Aspose.Cells | fit worksheet width PDF conversion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The sample loads an input.xlsx workbook, accesses the first worksheet, configures PageSetup.FitToPagesWide = 1 and FitToPagesTall = 0 to compress all columns into one PDF page width, optionally sets the print area to the used range, applies PdfSaveOptions with OnePagePerSheet = false, and saves the result as output.pdf while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Configure page setup to fit all columns on a single PDF page
            PageSetup pageSetup = sheet.PageSetup;
            pageSetup.FitToPagesWide = 1;   // Fit all columns to one page width
            pageSetup.FitToPagesTall = 0;   // No restriction on rows (allow multiple pages vertically)

            // Optional: set the print area to the used range to avoid empty cells
            // Use RefersTo to obtain the address string of the range
            pageSetup.PrintArea = sheet.Cells.MaxDisplayRange.RefersTo;

            // Prepare PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Ensure each worksheet is rendered according to its page setup
                OnePagePerSheet = false
            };

            // Save the workbook as a PDF with the configured page setup
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
