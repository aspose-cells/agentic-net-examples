// Title: Convert an Excel workbook to PDF with exact dimensions by disabling page scaling in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, configures the first worksheet's PageSetup to 100 % zoom (IsPercentScale = true, Zoom = 100), and saves the workbook as a PDF. | Demonstrate how to export an Excel workbook to PDF with Aspose.Cells while preserving the original page size by turning off percent scaling in the export options.
// Common Searches: asp.net aspocells export excel to pdf without changing page size | c# set worksheet page setup zoom 100 before pdf conversion aspocells | how to keep original dimensions when converting xlsx to pdf using aspocells | pdfsaveoptions default scaling aspocells c# example
// Tags: Aspose.Cells workbook to PDF without scaling | disable percent scaling in PageSetup | Worksheet zoom 100 export PDF | PdfSaveOptions default scaling Aspose.Cells | maintain page size during Excel to PDF conversion

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input Excel file, loads it into an Aspose.Cells Workbook, sets the first worksheet's PageSetup to use a 100 % zoom (IsPercentScale = true, Zoom = 100) to prevent any scaling, creates default PdfSaveOptions, and saves the workbook as a PDF while handling potential exceptions.
class WorkbookToPdf
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Set 100% zoom (no scaling) for the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.PageSetup.IsPercentScale = true;
            sheet.PageSetup.Zoom = 100;

            // Use default PDF save options (preserve original size)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
