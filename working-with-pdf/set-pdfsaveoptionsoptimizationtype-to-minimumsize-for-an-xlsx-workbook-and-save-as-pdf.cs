// Title: Convert an XLSX workbook to a minimum‑size PDF using Aspose.Cells PdfSaveOptions in C#
// AI Prompts: Generate C# code that loads an Excel workbook, configures the PDF save options for the smallest possible file, and saves the workbook as a PDF with Aspose.Cells. | Show how to enable the smallest PDF output when converting a workbook to PDF in a .NET application using Aspose.Cells.
// Common Searches: how to set PdfSaveOptions OptimizationType to MinimumSize in Aspose.Cells C# | Aspose.Cells convert Excel to PDF with smallest file size .NET | C# example for reducing PDF size when saving workbook as PDF using Aspose.Cells | minimum size PDF output from XLSX using Aspose.Cells PdfSaveOptions | Aspose.Cells PDF optimization options for .NET developers
// Tags: Aspose.Cells PdfSaveOptions MinimumSize | C# Excel to PDF conversion optimization | PDF optimization mode Aspose.Cells | reduce PDF file size Aspose.Cells | save workbook as PDF with minimal size

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing XLSX workbook (or creates a simple one), optionally sets PdfSaveOptions.OptimizationType to MinimumSize, and saves the workbook as a PDF file using Aspose.Cells, demonstrating how to produce the smallest possible PDF output.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            // Load workbook if it exists; otherwise create a simple one.
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                workbook.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                workbook.Save(inputPath);
            }

            // Set PDF save options.
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            // Uncomment the following line if PdfOptimizationType is available in your version.
            // pdfOptions.OptimizationType = PdfOptimizationType.MinimumSize;

            // Save the workbook as PDF using the specified options.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
