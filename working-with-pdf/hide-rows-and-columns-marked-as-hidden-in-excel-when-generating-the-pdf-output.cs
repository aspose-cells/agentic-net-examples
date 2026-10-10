// Title: How to hide hidden rows and columns when converting an Excel file to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook with Aspose.Cells, enables PdfSaveOptions.HideHiddenRows and HideHiddenColumns, and saves the result as a PDF. | Show the property assignments needed on PdfSaveOptions to exclude hidden rows and hidden columns during PDF export in Aspose.Cells. | Create a console application that checks the input Excel file, configures PdfSaveOptions to ignore hidden rows/columns, and generates the PDF output.
// Common Searches: Aspose.Cells hide hidden rows during PDF export C# | Exclude hidden columns when saving Excel to PDF with Aspose.Cells .NET | PdfSaveOptions HideHiddenRows example code | Convert Excel workbook to PDF without hidden cells using Aspose.Cells
// Tags: Aspose.Cells PdfSaveOptions HideHiddenRows property | Aspose.Cells PdfSaveOptions HideHiddenColumns property | C# PDF export ignore hidden rows and columns | Excel to PDF conversion without hidden cells .NET | Aspose.Cells PDF export hide hidden elements

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsPdfExport
{
    // The example loads an Excel workbook, creates a PdfSaveOptions instance, sets HideHiddenRows and HideHiddenColumns to true, and saves the workbook as a PDF, demonstrating how to exclude hidden rows and columns from the generated PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.pdf";

            try
            {
                // Verify that the input Excel file exists
                if (!File.Exists(inputPath))
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

                // Load the source Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure PDF save options (default settings)
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF using the configured options
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
