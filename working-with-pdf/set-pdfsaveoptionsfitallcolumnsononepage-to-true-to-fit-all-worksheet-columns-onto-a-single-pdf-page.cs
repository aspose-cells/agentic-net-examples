// Title: How to fit all worksheet columns onto a single PDF page with Aspose.Cells PdfSaveOptions in C#
// Common Searches: Aspose.Cells C# export Excel to PDF with all columns on one page | Set PdfSaveOptions.FitAllColumnsOnOnePage true in .NET | C# code to render each worksheet as a single PDF page using Aspose.Cells | Fit all columns on one PDF page using Aspose.Cells PdfSaveOptions.OnePagePerSheet | Generate single-page PDF from an Excel workbook with Aspose.Cells in C# | Aspose.Cells PDFSaveOptions column scaling for single-page output
// Tags: Aspose.Cells PdfSaveOptions column fitting | C# export worksheet to single-page PDF | OnePagePerSheet PDF rendering option | FitAllColumnsOnOnePage Aspose.Cells | Excel to PDF single sheet layout C# | Aspose.Cells PDF column scaling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExample
{
    // The example creates a Workbook, adds sample data, configures PdfSaveOptions with OnePagePerSheet = true to render each worksheet on a single PDF page so all columns fit, and saves the result as a PDF file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook (or load an existing one)
                Workbook workbook;

                // Example of loading an existing workbook safely
                // string inputPath = "input.xlsx";
                // if (File.Exists(inputPath))
                // {
                //     workbook = new Workbook(inputPath);
                // }
                // else
                // {
                //     Console.WriteLine($"Input file not found: {inputPath}. Creating a new workbook instead.");
                //     workbook = new Workbook();
                // }

                // For this demo we create a new workbook directly
                workbook = new Workbook();

                // Add sample data to demonstrate column fitting
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Cells["A1"].PutValue("Header 1");
                sheet.Cells["B1"].PutValue("Header 2");
                sheet.Cells["C1"].PutValue("Header 3");
                sheet.Cells["A2"].PutValue("Data A");
                sheet.Cells["B2"].PutValue("Data B");
                sheet.Cells["C2"].PutValue("Data C");

                // Configure PDF save options to fit all columns on a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Ensures each worksheet is rendered on one PDF page (all columns fit)
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF using the configured options
                string outputPath = "output.pdf";
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
