// Title: Generate a minimalist PDF from an Excel workbook by ignoring cell background colors with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# export Excel to PDF without cell background colors | How to create a PDF from a workbook that omits shading using Aspose.Cells | C# PdfSaveOptions ignore cell fill Aspose.Cells example | Minimalist PDF conversion from .xlsx using Aspose.Cells .NET | Remove cell background when saving Excel as PDF with Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions ignore cell fill | C# export workbook to PDF minimal layout | Excel to PDF conversion without background colors | Aspose.Cells PDF conversion styling options | remove cell shading during PDF export

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file into an Aspose.Cells Workbook, configures PdfSaveOptions to suppress cell background colors (e.g., by setting IgnoreCellBackground = true), and saves the workbook as a PDF with a clean, minimalist appearance. It includes basic file existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (background colors are included by default)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // If you need to ignore cell borders, you can set:
            // pdfOptions.IgnoreCellBorder = true;

            // Save the workbook as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
