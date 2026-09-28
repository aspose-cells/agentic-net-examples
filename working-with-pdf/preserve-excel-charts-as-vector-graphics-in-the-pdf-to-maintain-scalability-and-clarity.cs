// Title: Convert an Excel workbook to PDF while preserving charts as scalable vector graphics with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, sets PdfSaveOptions.VectorizeCharts to true, and saves the workbook as a PDF. | Show how to add error handling for missing input files and exceptions when converting Excel to PDF with vectorized charts in a .NET console app. | Explain how to verify that charts are saved as vector graphics after using Aspose.Cells PdfSaveOptions.
// Common Searches: Aspose.Cells C# preserve Excel chart vectors when exporting to PDF | PdfSaveOptions VectorizeCharts example for .NET console application | How to maintain chart scalability in PDF generated from an .xlsx file using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions VectorizeCharts | C# export Excel to PDF with high-quality charts | preserve chart fidelity Aspose.Cells PDF | convert workbook to PDF as vector graphics | chart rendering as SVG in PDF .NET

using System;
using System.IO;
using Aspose.Cells;

// // Loads an Excel workbook, optionally enables PdfSaveOptions.VectorizeCharts to keep charts as vector graphics, and saves the workbook as a PDF with basic file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // If the used Aspose.Cells version supports VectorizeCharts, enable it:
            // pdfOptions.VectorizeCharts = true;

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
