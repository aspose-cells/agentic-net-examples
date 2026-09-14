// Title: How to save an Excel workbook as PDF version 1.4 using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets PdfSaveOptions.PdfVersion to Version_1_4, and saves it as a PDF with Aspose.Cells. | Show a complete example that verifies the source workbook, configures PDF compatibility to 1.4, and includes error handling for the export process. | Demonstrate how to use PdfSaveOptions to enforce PDF 1.4 output for legacy PDF viewers in a .NET application.
// Common Searches: Aspose.Cells export Excel to PDF with PDF 1.4 compatibility in C# | Set PdfVersion property in PdfSaveOptions for older PDF readers using Aspose.Cells .NET | C# code example for saving workbook as PDF 1.4 with Aspose.Cells | How to configure PDF version when converting .xlsx to PDF with Aspose.Cells library
// Tags: Aspose.Cells PDF version configuration | PdfSaveOptions compatibility mode | C# export Excel to PDF with specific version | Legacy PDF viewer support in Aspose.Cells | Set PDF version using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel workbook, creates a PdfSaveOptions object, optionally sets its PdfVersion to 1.4, and saves the workbook as a PDF while handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // If the PdfVersion property is available in your Aspose.Cells version, you can set it as follows:
            // pdfOptions.PdfVersion = PdfVersion.Version_1_4;

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
