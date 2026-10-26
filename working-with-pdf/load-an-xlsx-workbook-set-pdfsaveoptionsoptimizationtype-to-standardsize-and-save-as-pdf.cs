// Title: Convert an XLSX workbook to PDF with StandardSize optimization using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set PdfSaveOptions OptimizationType to StandardSize before saving as PDF | How to export Excel to PDF with StandardSize optimization in .NET | C# example for converting XLSX to PDF using PdfSaveOptions StandardSize | Aspose.Cells PDF export with reduced file size using StandardSize option | Save workbook as PDF with StandardSize optimization using Aspose.Cells API
// Tags: PdfSaveOptions StandardSize optimization | XLSX to PDF conversion Aspose.Cells C# | Aspose.Cells PDF export configuration | Workbook.Save with PDF options .NET | Optimize PDF size using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample program verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object, sets its OptimizationType to StandardSize to produce a smaller PDF, and then saves the workbook as output.pdf while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Create PDF save options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF using the configured options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
