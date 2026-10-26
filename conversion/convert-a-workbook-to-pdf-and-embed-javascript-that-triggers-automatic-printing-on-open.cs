// Title: Convert an Excel workbook to PDF using Aspose.Cells for .NET (C#) with robust error handling
// Common Searches: Aspose.Cells C# convert .xlsx to PDF with PdfSaveOptions example | C# verify Excel file exists before saving as PDF using Aspose.Cells | How to handle exceptions when exporting Excel to PDF in .NET with Aspose.Cells | Programmatic Excel to PDF conversion code sample for Aspose.Cells in C# | Save workbook as PDF in C# using Aspose.Cells and check for missing input file
// Tags: Aspose.Cells PDF conversion C# | PdfSaveOptions Aspose.Cells | Excel to PDF export .NET | file existence validation Aspose.Cells | exception handling Aspose.Cells PDF save

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the presence of an input.xlsx file, loads it into an Aspose.Cells Workbook, configures PdfSaveOptions, and saves the workbook as output.pdf while gracefully handling any runtime errors.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Initialize PDF save options (no constructor arguments needed)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
