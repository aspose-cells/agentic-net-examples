// Title: C# – Convert an XLSB workbook to PDF with Aspose.Cells, including file check and error handling
// Common Searches: how to export an XLSB workbook to PDF using Aspose.Cells in C# | Aspose.Cells C# example for converting binary Excel file to PDF with try‑catch | C# code to verify XLSB file exists before converting to PDF with Aspose.Cells | save a workbook as PDF in .NET using Aspose.Cells SaveFormat.Pdf | convert large XLSB to PDF in C# while handling missing file errors
// Tags: Aspose.Cells PDF export from binary Excel | C# workbook.Save with SaveFormat.Pdf | file existence validation for Aspose.Cells workbook | exception handling for Aspose.Cells PDF conversion | binary Excel workbook conversion in .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks that the specified XLSB file is present, loads it into an Aspose.Cells Workbook, and saves the workbook as a PDF using SaveFormat.Pdf. It includes try‑catch error handling to report any issues during the conversion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as a PDF document
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF file successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
