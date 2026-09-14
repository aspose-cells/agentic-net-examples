// Title: How to convert a macro-enabled XLSM workbook to PDF in C# using Aspose.Cells
// AI Prompts: Generate C# code that checks for an .xlsm file, loads it with Aspose.Cells Workbook, and saves it as a PDF with proper exception handling. | Show a console application example that validates the input path before converting a macro-enabled Excel file to PDF using Aspose.Cells. | Provide a minimal Aspose.Cells workflow in C# for reading an XLSM workbook and exporting it to PDF format.
// Common Searches: c# aspnet convert macro enabled xlsm to pdf using aspose.cells library | how to verify excel file exists before saving as pdf with aspose.cells | asp.net console app export xlsm workbook to pdf with error handling | aspose.cells save workbook as pdf without losing macros | sample code for converting xlsm to pdf in .NET
// Tags: aspose.cells xlsm to pdf conversion | c# workbook save as pdf with aspose.cells | file existence check before aspose.cells export | asp.net console pdf export from macro-enabled excel | exception handling for aspose.cells pdf save

using System;
using System.IO;
using Aspose.Cells;

// A C# console program that verifies the presence of an input.xlsm file, loads it into an Aspose.Cells Workbook, and saves the workbook as output.pdf using the Pdf SaveFormat, while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsm";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the XLSM workbook
            Workbook workbook = new Workbook(inputPath);

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook successfully saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
