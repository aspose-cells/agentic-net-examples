// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET (PDF language setting not supported)
// Common Searches: how to convert xlsx to pdf with Aspose.Cells in C# | Aspose.Cells PDF conversion example C# | save Excel workbook as PDF using Aspose.Cells .NET | does Aspose.Cells allow setting PDF document language during conversion | C# PdfSaveOptions language property missing in Aspose.Cells
// Tags: Aspose.Cells Excel to PDF conversion C# | PdfSaveOptions configuration Aspose.Cells | PDF document language not supported Aspose.Cells | file existence check before conversion C# | exception handling Aspose.Cells workbook save

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the input XLSX file, loads it into an Aspose.Cells Workbook, creates a PdfSaveOptions object (which lacks a language property), and saves the workbook as a PDF. It demonstrates basic error handling and console output confirming the conversion, while noting that PDF language cannot be set via Aspose.Cells.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the source Excel workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (language setting not supported in PdfSaveOptions)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Convert and save the workbook as a PDF file using the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
