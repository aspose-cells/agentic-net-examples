// Title: Export an XLS workbook containing charts to PDF with Aspose.Cells for .NET while preserving chart appearance
// Common Searches: Aspose.Cells C# export XLS file with charts to PDF preserving chart layout | how to save Excel .xls workbook as PDF with charts using Aspose.Cells .NET | C# convert legacy Excel workbook to PDF and keep chart rendering | PdfSaveOptions OnePagePerSheet false chart rendering Aspose.Cells example | export Excel file with embedded charts to PDF using Aspose.Cells API
// Tags: Aspose.Cells XLS to PDF with chart rendering | PdfSaveOptions OnePagePerSheet false | C# export Excel charts to PDF | preserve Excel chart appearance Aspose.Cells | convert legacy .xls workbook to PDF

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the input XLS file, loads it with Aspose.Cells, configures PdfSaveOptions (including OnePagePerSheet = false), and saves the workbook as a PDF. All embedded charts are rendered in the output PDF, and the code includes basic error handling for missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the XLS workbook that contains charts
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (charts are rendered by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example option: keep original page layout
                OnePagePerSheet = false
            };

            // Export the workbook to PDF, preserving chart rendering
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
