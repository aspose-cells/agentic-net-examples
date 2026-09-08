// Title: Convert an XLSB workbook to PDF using Aspose.Cells for .NET with file existence verification and error handling
// AI Prompts: Generate C# code that loads an .xlsb file with Aspose.Cells, checks that the file exists, and saves the workbook as a PDF using PdfSaveOptions. | Create a console application in .NET that catches and logs any exceptions while converting an Excel binary workbook to PDF.
// Common Searches: asp.net convert xlsb to pdf with aspose.cells and verify file existence | c# load xlsb workbook and save as pdf using pdfsaveoptions | handle errors when converting excel binary files to pdf in c#
// Tags: aspose.cells xlsb to pdf conversion c# | c# verify excel file existence before conversion | pdfsaveoptions error handling asp.net

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an XLSB workbook with Aspose.Cells, confirming the source file is present, and saving the workbook as a PDF using PdfSaveOptions, while gracefully handling any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsb";
        const string outputPath = "output.pdf";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the XLSB workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options (default options are sufficient here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);

            Console.WriteLine($"Workbook successfully saved as PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
