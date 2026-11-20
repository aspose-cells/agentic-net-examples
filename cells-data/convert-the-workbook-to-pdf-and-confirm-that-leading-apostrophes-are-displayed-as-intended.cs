// Title: Convert an Excel workbook to PDF while preserving leading apostrophes using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# convert .xlsx to PDF keep leading apostrophe | How to retain leading apostrophe characters when saving Excel as PDF with Aspose.Cells | C# PDF conversion of Excel workbook preserving cell text that starts with an apostrophe | Save workbook as PDF with Aspose.Cells and display leading apostrophes correctly | Aspose.Cells PdfSaveOptions default handling of leading apostrophes in PDF output
// Tags: Aspose.Cells PDF conversion preserving leading apostrophes | PdfSaveOptions default apostrophe display | Excel to PDF conversion C# Aspose.Cells | Workbook.Save PDF with leading apostrophe handling | Apostrophe-leading cell values retained in PDF output

using System;
using System.IO;
using Aspose.Cells;

// Loads 'input.xlsx', uses the default PdfSaveOptions, and saves it as 'output.pdf' ensuring that any leading apostrophes in cell values are displayed correctly in the generated PDF.
class Program
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
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default behavior displays leading apostrophes)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine("Workbook successfully converted to PDF with leading apostrophes displayed.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
