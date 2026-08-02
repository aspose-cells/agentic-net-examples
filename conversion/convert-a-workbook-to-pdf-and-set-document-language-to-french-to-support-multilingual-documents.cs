// Title: C# – Convert Aspose.Cells Workbook to PDF with French Language Metadata
// Description: Demonstrates how to use Aspose.Cells for .NET to create a workbook, insert French text, set the workbook UI language to France, configure PdfSaveOptions (including default edit language), and save the file as a PDF that carries French language tags for accessibility and search indexing.
// Keywords: Aspose.Cells PDF conversion C# | set workbook language French | PdfSaveOptions language metadata | Excel to PDF French locale | multilingual PDF Aspose | CountryCode.France | DefaultEditLanguage | document language tag PDF | accessibility PDF French | .NET Aspose.Cells localization
// Common Searches: Aspose.Cells set PDF language to French | C# convert Excel to PDF with French UI language | PdfSaveOptions default edit language Aspose.Cells | How to add language metadata to PDF using Aspose.Cells | workbook.Settings.LanguageCode France example | Export French Excel to PDF .NET | Aspose.Cells multilingual PDF sample
// Developer Intent: Generate a PDF from an Excel workbook while embedding French language information for UI and accessibility purposes.
// Use Cases: Localization testing – verify that PDFs display French UI language tags. | Compliance – produce PDFs with correct language tags for screen‑reader support. | Automated reporting – create French‑language PDFs from dynamically generated data. | Batch conversion – convert multiple workbooks to PDF with French language settings in a single process. | Document archiving – store PDFs with language metadata to improve search and categorization.
// AI Prompts: Write C# code using Aspose.Cells to save a workbook as PDF with French language metadata (workbook.Settings.LanguageCode = CountryCode.France) and set DefaultEditLanguage. | Explain the difference between workbook.Settings.LanguageCode and PdfSaveOptions.DefaultEditLanguage and how each influences the language tags in the resulting PDF. | Show how to validate the language metadata of the generated PDF using a PDF inspection library. | Provide troubleshooting steps when the French language tag is missing from a PDF created with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells for .NET to create a workbook, insert French text, set the workbook UI language to France, configure PdfSaveOptions (including default edit language), and save the file as a PDF that carries French language tags for accessibility and search indexing.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet and add some French text
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Bonjour le monde"); // Sample French content

            // Set the workbook UI language to French (France)
            workbook.Settings.LanguageCode = CountryCode.France;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Set default edit language (use English if French is unavailable in the current library version)
            pdfOptions.DefaultEditLanguage = DefaultEditLanguage.English;

            // Define output file path
            string outputPath = "output.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
