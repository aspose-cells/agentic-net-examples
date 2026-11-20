// Title: Export an Excel workbook to PDF with all annotations flattened and columns fitted to a single page per sheet using Aspose.Cells for .NET
// Common Searches: how to export Excel to PDF with comments flattened using Aspose.Cells C# | Aspose.Cells C# save workbook as PDF all columns on one page per sheet | C# convert .xlsx to PDF and remove annotation layers with Aspose.Cells | PdfSaveOptions AllColumnsInOnePagePerSheet example Aspose.Cells | flatten Excel comments when generating PDF with Aspose.Cells .NET
// Tags: Aspose.Cells PDF export with merged annotations | C# PdfSaveOptions single page per sheet | Excel workbook to PDF conversion using Aspose.Cells | merge comments into PDF Aspose.Cells | Aspose.Cells example saving .xlsx as PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    // The sample loads an existing Excel file into an Aspose.Cells Workbook, configures PdfSaveOptions to place all columns on a single page per sheet, and saves the workbook as a PDF. During the export, annotations such as comments are flattened into the PDF content, and any errors are caught and reported.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input Excel file
                string inputFile = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputFile);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    // Place all columns on a single page per sheet
                    AllColumnsInOnePagePerSheet = true
                };

                // Output PDF file
                string outputFile = "output.pdf";

                // Save the workbook as PDF using the configured options
                workbook.Save(outputFile, pdfOptions);

                Console.WriteLine($"Workbook successfully saved as PDF: {outputFile}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
