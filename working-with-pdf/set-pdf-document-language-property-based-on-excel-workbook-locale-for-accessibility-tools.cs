// Title: Set the PDF language property from an Excel workbook’s locale using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# set PDF document language from workbook locale | How to add language metadata to PDF generated from Excel with Aspose.Cells | C# PdfSaveOptions language property for accessibility when converting Excel to PDF | Assign PDF language tag based on Excel workbook culture using Aspose.Cells | Define PDF accessibility language during Excel to PDF conversion in C#
// Tags: Aspose.Cells PdfSaveOptions language property | Excel workbook locale to PDF language | PDF accessibility metadata C# | Convert Excel to PDF with language tag | Set PDF document language programmatically

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPdfExample
{
    // The sample loads an Excel workbook with Aspose.Cells, creates a PdfSaveOptions object, and saves the workbook as a PDF. By setting the PdfSaveOptions.Language property to the workbook's locale, you can embed the appropriate language metadata for accessibility tools.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                    return;
                }

                // Load the Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no additional PDF document info required)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
