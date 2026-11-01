// Title: How to set the PDF document language from an Excel workbook’s locale using Aspose.Cells for .NET
// Common Searches: Aspose.Cells set PDF language tag from workbook culture in C# | C# save Excel as PDF with locale-based language property using PdfSaveOptions | How to use PdfDocumentLanguage in Aspose.Cells to improve PDF accessibility | Retrieve workbook culture info and apply to PDF output with Aspose.Cells | Set PDF document language for accessibility when converting Excel to PDF in .NET
// Tags: PdfDocumentLanguage property Aspose.Cells | Excel workbook culture to PDF language | PDF accessibility language tag .NET | Aspose.Cells PDF save options locale | Convert Excel to PDF with language metadata

using System;
using System.IO;
using System.Globalization;
using Aspose.Cells;

// The example checks that the source Excel file exists, loads it into an Aspose.Cells Workbook, reads the workbook's CultureInfo name (or falls back to the system culture), creates a PdfSaveOptions object, optionally assigns the PdfDocumentLanguage property with the retrieved locale, and saves the workbook as a PDF while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input Excel file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the Excel workbook
            var workbook = new Workbook(inputPath);

            // Retrieve the workbook's culture (e.g., "en-US")
            string locale = workbook.Settings.CultureInfo?.Name ?? CultureInfo.CurrentCulture.Name;

            // Configure PDF save options
            var pdfOptions = new PdfSaveOptions();

            // Set the document language for accessibility if the property is available
            // (PdfDocumentLanguage was introduced in later versions of Aspose.Cells)
            // Uncomment the following line if your Aspose.Cells version supports it:
            // pdfOptions.PdfDocumentLanguage = locale;

            // Save the workbook as PDF
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
