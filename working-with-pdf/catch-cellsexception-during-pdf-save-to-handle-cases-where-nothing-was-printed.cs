// Title: C# – Catch CellsException for "nothing was printed" when saving an Excel workbook to PDF with Aspose.Cells
// Common Searches: Aspose.Cells C# catch CellsException when PDF export reports nothing was printed | handle no printable content error during Excel to PDF conversion using Aspose.Cells | detect empty worksheet and skip PDF generation with Aspose.Cells in .NET | C# try‑catch CellsException for empty workbook PDF save Aspose.Cells | how to check for "nothing was printed" message in CellsException while saving PDF
// Tags: Aspose.Cells PDF save exception handling | CellsException handling in C# | detect empty worksheet during PDF export | Excel to PDF conversion error handling | Aspose.Cells no printable content scenario

using Aspose.Cells;
using Aspose.Cells.Rendering;
using System;
using System.IO;

// The example loads an Excel file, attempts to save it as a PDF using Aspose.Cells, and catches CellsException to identify the "nothing was printed" condition. It logs a friendly message when the workbook contains no printable content, rethrows other CellsExceptions, and also handles unexpected errors during PDF conversion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists before attempting to load it.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        Workbook workbook;
        try
        {
            // Load the workbook from the specified file.
            workbook = new Workbook(inputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to load workbook: {ex.Message}");
            return;
        }

        // Configure PDF save options if required.
        PdfSaveOptions pdfOptions = new PdfSaveOptions();
        // Example option: one page per sheet
        // pdfOptions.OnePagePerSheet = true;

        try
        {
            // Save the workbook as PDF using the options object.
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells may throw a CellsException when there is nothing to print.
            bool noDataToPrint = false;

            // Fallback: inspect the exception message for known text.
            if (ex.Message != null && ex.Message.IndexOf("nothing was printed", StringComparison.OrdinalIgnoreCase) >= 0)
                noDataToPrint = true;

            if (noDataToPrint)
            {
                // Handle the "nothing was printed" scenario gracefully.
                Console.WriteLine("PDF generation skipped: no printable content found in the workbook.");
            }
            else
            {
                // Re‑throw or handle other CellsExceptions as appropriate.
                Console.WriteLine($"CellsException occurred: {ex.Message}");
                throw;
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions.
            Console.WriteLine($"An error occurred while saving PDF: {ex.Message}");
        }
    }
}
