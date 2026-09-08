// Title: Generate a PDF from an Excel workbook in C# while automatically excluding hidden rows and columns using Aspose.Cells
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, configures PdfSaveOptions to ignore hidden rows and columns, and saves the workbook as a PDF. | Show how to check for the existence of the source Excel file, handle a missing‑file error, and convert it to PDF while preserving the hidden state using Aspose.Cells. | Demonstrate setting up Aspose.Cells PdfSaveOptions in C# so that rows or columns marked as hidden in the workbook are not rendered in the resulting PDF.
// Common Searches: how to hide hidden rows and columns when converting Excel to PDF with Aspose.Cells in C# | Aspose.Cells PdfSaveOptions exclude hidden rows .NET example | C# convert .xlsx to PDF without showing hidden columns using Aspose.Cells | skip hidden rows during Excel to PDF conversion with Aspose.Cells library
// Tags: Aspose.Cells PdfSaveOptions hide rows columns | C# Excel to PDF conversion excluding hidden elements | Aspose.Cells workbook.Save PDF hidden rows | error handling missing Excel file Aspose.Cells | convert .xlsx to PDF respecting hidden state

using System;
using System.IO;
using Aspose.Cells;

// The example verifies that the input Excel file exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which by default omits hidden rows and columns), and saves the workbook as a PDF while handling any runtime exceptions.
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
            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default behavior hides hidden rows/columns)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF file using the configured options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
