// Title: Export an Excel workbook to a 300 DPI high‑resolution PDF using Aspose.Cells for .NET
// Common Searches: how to export Excel to PDF with 300 DPI using Aspose.Cells C# | Aspose.Cells set PDF image DPI to 300 in .NET | C# generate high‑resolution PDF from .xlsx with Aspose.Cells | increase PDF export quality to 300 DPI with Aspose.Cells PdfSaveOptions | save workbook as high‑resolution PDF Aspose.Cells .NET example
// Tags: Aspose.Cells PdfSaveOptions ImageDPI | export workbook to high‑resolution PDF C# | set PDF export DPI Aspose.Cells | generate 300 DPI PDF from Excel .NET | high quality PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The code checks that 'input.xlsx' exists, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object with ImageDPI set to 300, and saves the workbook as 'output.pdf'. Exception handling reports any errors that occur during the process.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF export options (default settings)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
