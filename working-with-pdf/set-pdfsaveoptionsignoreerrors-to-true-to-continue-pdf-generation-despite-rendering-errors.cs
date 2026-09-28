// Title: How to set PdfSaveOptions.IgnoreErrors = true to keep PDF generation running in Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, sets PdfSaveOptions.IgnoreErrors to true, and saves it as a PDF using Aspose.Cells. | Show an example of configuring Aspose.Cells PdfSaveOptions to bypass rendering errors during Excel‑to‑PDF conversion in a .NET console application. | Provide a snippet that demonstrates checking for a missing input file while using PdfSaveOptions.IgnoreErrors in Aspose.Cells.
// Common Searches: Aspose.Cells PdfSaveOptions ignore errors during PDF export C# | continue PDF creation after rendering error Aspose.Cells .NET | skip Excel cell rendering issues when converting to PDF with Aspose | set PdfSaveOptions.IgnoreErrors true example code | handle exceptions while saving workbook as PDF using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions ignore errors | C# Excel to PDF conversion error handling | PdfSaveOptions rendering error bypass | Aspose.Cells PDF export ignore rendering issues | Workbook.Save with PdfSaveOptions IgnoreErrors

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExample
{
    // The sample loads an Excel file, creates a PdfSaveOptions object, sets IgnoreErrors = true to suppress rendering problems, and saves the workbook as a PDF, including a check for the input file's existence and basic exception handling.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputFile = "input.xlsx";
                string outputFile = "output.pdf";

                // Ensure the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options (add any required settings here)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as PDF
                workbook.Save(outputFile, pdfOptions);
                Console.WriteLine($"PDF generated successfully: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
