// Title: Create a landscape-oriented PDF from a wide Excel workbook with all columns on one page using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert wide Excel sheet to landscape PDF with all columns on a single page | How to set worksheet orientation to landscape when saving Excel as PDF with Aspose.Cells | Fit all columns onto one PDF page per sheet using PdfSaveOptions in Aspose.Cells | C# example for exporting Excel to PDF in landscape mode with Aspose.Cells | OnePagePerSheet and AllColumnsInOnePagePerSheet settings for PDF conversion in Aspose.Cells
// Tags: Aspose.Cells landscape PDF conversion | PdfSaveOptions OnePagePerSheet C# | AllColumnsInOnePagePerSheet Aspose.Cells | set worksheet orientation landscape Aspose.Cells | export wide Excel sheet to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing XLSX file, sets each worksheet's page orientation to landscape, configures PdfSaveOptions to fit all columns onto a single PDF page per sheet, and saves the workbook as a PDF while handling errors.
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

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Set landscape orientation for each worksheet (required when DefaultPageSetup is unavailable)
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.PageSetup.Orientation = PageOrientationType.Landscape;
            }

            // Create PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Fit all columns of each sheet onto a single PDF page width
                OnePagePerSheet = true,
                AllColumnsInOnePagePerSheet = true
            };

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: \"{outputPath}\"");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
