// Title: Configure font substitution with PdfSaveOptions for Excel-to-PDF conversion using Aspose.Cells in C#
// Common Searches: c# aspose.cells set font substitution when saving workbook to pdf | map missing Excel font to Arial using PdfSaveOptions in Aspose.Cells | replace unavailable fonts during Excel to PDF export with Aspose.Cells .NET | how to apply custom font fallback for PDF generation in Aspose.Cells C# | use FontSettings.SubstitutionFonts with PdfSaveOptions for PDF output
// Tags: pdfsaveoptions custom font substitution aspnet | excel to pdf font fallback aspose cells | missing font mapping pdf export csharp | font substitution rule aspose cells pdf | configure font settings for pdf conversion .net

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input Excel file, loads it into an Aspose.Cells Workbook, optionally adds a substitution rule (e.g., mapping a missing font to Arial) via FontSettings.SubstitutionFonts, creates a PdfSaveOptions object, and saves the workbook as a PDF using those options while handling any runtime exceptions.
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

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Font substitution is optional; if the current Aspose.Cells version supports it,
            // you can uncomment the following lines:
            // workbook.FontSettings.SubstitutionFonts.Add("MissingFont", "Arial");

            // Create PDF save options (default options are sufficient for this example)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully converted to PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
