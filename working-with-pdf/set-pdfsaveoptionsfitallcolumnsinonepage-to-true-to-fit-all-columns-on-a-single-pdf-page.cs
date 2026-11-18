// Title: How to fit all worksheet columns onto a single PDF page using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF with all columns on one page | fit worksheet columns to a single PDF page using Aspose.Cells PdfSaveOptions | C# set FitAllColumnsInOnePage true in Aspose.Cells PDF conversion | configure page setup for one-page-wide PDF output in Aspose.Cells | Aspose.Cells PDFSaveOptions OnePagePerSheet and column fitting example
// Tags: Aspose.Cells PDF column fitting | PdfSaveOptions FitAllColumnsInOnePage property | C# Excel to PDF single-page layout | Worksheet PageSetup FitToPagesWide configuration | OnePagePerSheet PDF export Aspose.Cells | fit columns on one PDF page .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The sample checks for the existence of an input Excel file, loads it into an Aspose.Cells Workbook, and configures each worksheet's PageSetup to fit all columns within one page width (FitToPagesWide = 1, FitToPagesTall = 0). It then creates a PdfSaveOptions object with OnePagePerSheet enabled, saves the workbook as a PDF, and gracefully handles any runtime exceptions.
    class Program
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

                // Set each worksheet to fit all columns on one page
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Fit to 1 page wide, unlimited pages tall (0)
                    sheet.PageSetup.FitToPagesWide = 1;
                    sheet.PageSetup.FitToPagesTall = 0;
                }

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Optional: save each worksheet on a separate page
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
