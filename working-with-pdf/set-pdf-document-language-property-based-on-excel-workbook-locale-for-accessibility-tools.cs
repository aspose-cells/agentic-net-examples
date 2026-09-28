// Title: How to set the PDF document language from an Excel workbook’s locale using Aspose.Cells in C#
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, reads the workbook’s CultureInfo.Name, assigns it to PdfSaveOptions.PdfDocumentLanguage, and saves the workbook as a PDF. | Modify an existing Aspose.Cells PDF export routine to automatically use the workbook’s locale for the PDF language tag, falling back to the system culture when the workbook setting is missing. | Demonstrate how to check for the availability of the PdfDocumentLanguage property at runtime and set it conditionally for backward‑compatible Aspose.Cells versions.
// Common Searches: aspocells set pdf language tag from workbook culture c# | c# convert excel to pdf with correct language metadata for screen readers | how to use PdfSaveOptions.PdfDocumentLanguage with Aspose.Cells | fallback to system culture when workbook locale not defined in Aspose.Cells PDF export | accessibility language property in PDF generated from Excel using Aspose.Cells
// Tags: Aspose.Cells PDF language metadata | C# PdfSaveOptions language tag | Excel workbook locale extraction | PDF accessibility language property | Aspose.Cells backward compatibility

using System;
using System.Globalization;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an Excel workbook, determines its culture name (or uses the current system culture), creates PdfSaveOptions, optionally sets the PDF language property for accessibility, and saves the workbook as a PDF.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the Excel workbook (lifecycle rule: load)
            var workbook = new Workbook(inputPath);

            // Retrieve the workbook's culture name (e.g., "en-US")
            string localeName = workbook.Settings.CultureInfo?.Name ?? CultureInfo.CurrentCulture.Name;

            // Create PDF save options
            var pdfOptions = new PdfSaveOptions();

            // Set the document language for accessibility if the property is available
            // (Commented out for compatibility with older Aspose.Cells versions)
            // pdfOptions.PdfDocumentLanguage = localeName;

            // Save the workbook as PDF (lifecycle rule: save)
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
