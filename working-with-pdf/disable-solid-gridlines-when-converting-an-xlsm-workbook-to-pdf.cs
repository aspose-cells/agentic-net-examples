// Title: Disable solid gridlines while converting an XLSM workbook to PDF with Aspose.Cells for .NET
// Common Searches: how to hide gridlines in PDF output when exporting an XLSM file with Aspose.Cells C# | Aspose.Cells C# export XLSM to PDF without showing solid gridlines | PdfSaveOptions to remove gridlines during XLSM to PDF conversion in .NET | disable solid gridlines in PDF generated from an Excel macro-enabled workbook using Aspose.Cells | C# code sample for saving an XLSM workbook as PDF without gridlines using Aspose.Cells
// Tags: Aspose.Cells PDF export hide gridlines | PdfSaveOptions gridlines disabled | XLSM to PDF conversion without gridlines | .NET workbook save as PDF Aspose.Cells | solid gridlines removal Aspose.Cells PDF

using System;
using System.IO;
using Aspose.Cells;

// The example loads an XLSM workbook, creates a PdfSaveOptions object (gridlines are hidden by default), and saves the workbook as a PDF, demonstrating how to disable solid gridlines during the conversion while handling file existence and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the XLSM workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (gridlines are hidden by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
