// Title: Convert only the first worksheet of an Excel file to a PDF while preserving SVG graphics as vectors using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx workbook, extracts the first worksheet, and saves it as a PDF with one page per sheet using Aspose.Cells. | Adjust the conversion so that any SVG images embedded in the worksheet remain vector elements in the generated PDF. | Add comprehensive error handling that verifies the source file exists, catches conversion exceptions, and logs the output PDF path on success.
// Common Searches: Aspose.Cells C# export first worksheet to PDF preserving SVG as vector | keep SVG graphics vectorized when converting Excel to PDF with Aspose.Cells | C# PdfSaveOptions OnePagePerSheet example for single‑sheet workbook | validate input file before converting Excel to PDF using Aspose.Cells .NET
// Tags: export first worksheet to PDF Aspose.Cells | preserve SVG vector graphics in PDF conversion | PdfSaveOptions OnePagePerSheet setting | C# copy worksheet Aspose.Cells | file existence validation Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks that input.xlsx exists, loads it with Aspose.Cells, creates a new workbook containing only the first worksheet, configures PdfSaveOptions with OnePagePerSheet=true, and saves the result as output.pdf while handling and reporting any errors.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.pdf";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    throw new FileNotFoundException($"The input file '{inputPath}' was not found.");
                }

                // Load the original Excel workbook
                Workbook sourceWorkbook = new Workbook(inputPath);

                // Create a new workbook that will contain only the first worksheet
                Workbook singleSheetWorkbook = new Workbook();

                // Remove the default sheet created with the new workbook
                singleSheetWorkbook.Worksheets.RemoveAt(0);

                // Copy the first worksheet from the source workbook by name
                singleSheetWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0].Name);

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the single‑sheet workbook as PDF using the options
                singleSheetWorkbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF file successfully created at '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log or display the error details
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
