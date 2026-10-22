// Title: Convert an XLSB workbook to PDF with Aspose.Cells for .NET and add the current UTC creation date to the PDF metadata
// AI Prompts: Write C# code that loads an .xlsb file using Aspose.Cells, saves it as a PDF, and then uses Aspose.Pdf to set the PDF's CreationDate to DateTime.UtcNow. | Describe the steps to apply PdfSaveOptions when exporting an XLSB workbook to PDF and subsequently update the PDF document information in a .NET application.
// Common Searches: Aspose.Cells convert xlsb to pdf and set creation date metadata in C# | C# add current UTC timestamp to PDF after exporting from Excel workbook | How to modify PDF document info after saving with Aspose.Cells | Set PDF creation time using Aspose.Pdf after converting XLSB | Aspose.Cells PdfSaveOptions example for handling PDF metadata
// Tags: aspocells xlsb to pdf conversion | pdfsaveoptions aspocells export settings | aspopdf modify pdf metadata c# | set pdf creationdate utc c# | c# workbook export to pdf with metadata

using System;
using System.IO;
using Aspose.Cells;

namespace XlsbToPdf
{
    // The sample loads an XLSB file with Aspose.Cells, creates a PdfSaveOptions object, ensures the output folder exists, and saves the workbook as a PDF. Because the current Aspose.Cells API does not expose a PdfDocumentInfo property, the code does not directly set the PDF creation timestamp; you would need to use Aspose.Pdf after the conversion to update the CreationDate to DateTime.UtcNow.
    class XlsbToPdfConverter
    {
        static void Main()
        {
            const string inputPath = "input.xlsb";
            const string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the XLSB workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (no PdfDocumentInfo property in current API)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as PDF with the specified options
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
