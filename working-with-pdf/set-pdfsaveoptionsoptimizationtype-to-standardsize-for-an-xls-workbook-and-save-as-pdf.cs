// Title: Convert an XLS workbook to PDF with Standard size optimization using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xls workbook with Aspose.Cells, configures PdfSaveOptions to use the Standard optimization mode, and exports the file as a PDF. | Demonstrate how to apply the Standard PDF optimization setting when converting an Excel workbook to PDF in a .NET application using Aspose.Cells.
// Common Searches: Aspose.Cells how to set PdfSaveOptions OptimizationType to Standard for Excel to PDF conversion | C# example saving .xls as .pdf with standard size optimization using Aspose.Cells | PdfSaveOptions Standard optimization type Aspose.Cells .NET code sample | Convert legacy XLS workbook to PDF with reduced file size using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions Standard mode | C# Excel to PDF conversion using Aspose.Cells | Apply Standard PDF optimization with Aspose.Cells | Export XLS workbook to PDF with size optimization | Aspose.Cells PDF export configuration .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Loads an existing XLS workbook, configures PdfSaveOptions with OptimizationType set to Standard, and saves the workbook as a PDF file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xls";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the existing XLS workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options with Standard optimization
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OptimizationType = PdfOptimizationType.Standard
            };

            // Save the workbook as a PDF using the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
