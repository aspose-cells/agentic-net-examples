// Title: Limit PDF export to the first 10 pages of an Excel workbook using Aspose.Cells PdfSaveOptions in C#
// Common Searches: C# Aspose.Cells limit PDF export to first 10 pages | How to set PdfSaveOptions.PageCount when saving Excel as PDF with Aspose.Cells | Export only the first ten pages of a workbook to PDF using Aspose.Cells .NET | Aspose.Cells PDF conversion page count restriction example C# | Saving Excel as PDF with a page count limit in Aspose.Cells
// Tags: PdfSaveOptions.PageCount | Aspose.Cells PDF page limit | C# Excel to PDF conversion with page count | Workbook.Save with PdfSaveOptions | password protected workbook handling Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, configures PdfSaveOptions with PageCount set to 10 to restrict the generated PDF to the first ten pages, and saves the file as PDF while handling missing files and password‑protected workbooks.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (no password supplied)
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (limit to first 10 pages)
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    PageCount = 10
                };

                // Save the workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to: {outputPath}");
            }
            catch (CellsException ex)
            {
                // Handle password‑protected files or other Cells exceptions
                if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    Console.WriteLine("The workbook is password protected. Provide a valid password to load it.");
                }
                else
                {
                    Console.WriteLine($"CellsException: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
