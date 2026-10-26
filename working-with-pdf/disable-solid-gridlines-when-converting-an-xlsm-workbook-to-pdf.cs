// Title: Disable gridlines when converting an XLSM workbook to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export XLSM to PDF without gridlines | How to hide Excel gridlines in PDF output with PdfSaveOptions | C# remove solid gridlines from PDF generated from a macro-enabled workbook | PdfSaveOptions set to hide gridlines in Aspose.Cells PDF conversion | Convert macro-enabled Excel file to PDF without showing gridlines in .NET
// Tags: Aspose.Cells PDFSaveOptions hide gridlines | XLSM to PDF conversion without gridlines | disable Excel gridlines in PDF export C# | macro workbook PDF export settings Aspose.Cells | gridlines off PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSM workbook, creates a PdfSaveOptions object (gridlines are disabled by default), and saves the workbook as a PDF, resulting in a PDF file that does not display solid gridlines.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the XLSM workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (gridlines are disabled by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
