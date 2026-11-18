// Title: How to add row and column headings to every page of a PDF generated from Excel using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export Excel to PDF with row and column headings on each page | Enable PrintHeadings when saving workbook as PDF using Aspose.Cells .NET | C# PdfSaveOptions include worksheet headings in PDF output Aspose.Cells | Print row numbers and column letters in PDF generated from Excel with Aspose.Cells | Set page setup to show headings for PDF conversion in Aspose.Cells C# example
// Tags: Aspose.Cells PDF export with worksheet headers | PdfSaveOptions include worksheet header | PageSetup API for PDF conversion | Excel to PDF conversion with column row labels Aspose | C# generate PDF preserving worksheet headings

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, enables the PageSetup.PrintHeadings flag for the first worksheet, configures PdfSaveOptions, and saves the workbook as a PDF where each page displays the row numbers and column letters as headings.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Load existing workbook or create a simple one if the file is missing.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header1");
                sheet.Cells["B1"].PutValue("Header2");
                sheet.Cells["A2"].PutValue("Data1");
                sheet.Cells["B2"].PutValue("Data2");
                // Save placeholder for future runs.
                workbook.Save(inputPath);
            }

            // Enable printing of row and column headings for the first worksheet.
            workbook.Worksheets[0].PageSetup.PrintHeadings = true;

            // Configure PDF save options (no PrintHeadings property needed here).
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
