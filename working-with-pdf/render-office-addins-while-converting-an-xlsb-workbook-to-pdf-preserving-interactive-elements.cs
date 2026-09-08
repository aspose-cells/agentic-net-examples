// Title: Convert an XLSB workbook to PDF with Aspose.Cells for .NET and handle missing file errors
// AI Prompts: Write a C# console program that checks whether an .xlsb file exists, loads it with Aspose.Cells, and saves it as a PDF using the default PdfSaveOptions, wrapping the whole process in a try‑catch block. | Show how to log detailed exception information when converting an XLSB workbook to PDF with Aspose.Cells in a .NET application. | Demonstrate customizing PdfSaveOptions (e.g., embedding fonts) while converting an XLSB file to PDF using Aspose.Cells in C#. | Provide a C# snippet that returns the output PDF path after successfully converting an XLSB workbook with Aspose.Cells.
// Common Searches: asp.net convert xlsb file to pdf using aspose.cells with file not found check | c# example for saving an xlsb workbook as pdf with Aspose.Cells PdfSaveOptions | how to handle exceptions when converting xlsb to pdf in a .NET console app | default pdf save options for Aspose.Cells workbook conversion | sample code to verify input file existence before Aspose.Cells PDF export
// Tags: xlsb to pdf conversion Aspose.Cells | Aspose.Cells PdfSaveOptions C# | file existence validation Aspose.Cells | exception handling workbook to pdf .NET | convert binary Excel workbook to PDF using Aspose | default PDF export settings Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example verifies that the source .xlsb file exists, loads it into an Aspose.Cells Workbook, and saves it as a PDF using the default PdfSaveOptions, while gracefully handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";
        const string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient for most scenarios)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
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
