// Title: Convert an XLSB workbook to PDF in C# with Aspose.Cells while handling the unavailable MinimizeSize setting
// AI Prompts: Generate C# code that checks for the presence of an XLSB file, loads it with Aspose.Cells, and saves it as a PDF, including a fallback when the MinimizeSize property is not supported. | Show how to export an XLSB workbook to a compressed PDF using Aspose.Cells in .NET, applying any alternative size‑reduction options and proper error handling.
// Common Searches: asp.net convert xlsb file to pdf using aspose.cells | c# reduce pdf file size when exporting xlsb with aspose.cells | how to check if xlsb exists before loading in aspose.cells c# | workaround for missing MinimizeSize property in Aspose.Cells PDF export | export binary Excel workbook to PDF with size optimization in .NET
// Tags: XLSB to PDF conversion Aspose.Cells | PDF size reduction Aspose.Cells | Aspose.Cells workbook loading C# | missing MinimizeSize property handling | file existence validation C# Aspose.Cells | alternative PDF compression options Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates C# code that verifies an XLSB file exists, loads it with Aspose.Cells, notes the MinimizeSize option is unavailable, and saves the workbook as a PDF with basic exception handling and a placeholder for alternative size‑optimization techniques.
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

            // NOTE: The MinimizeSize property is not available in the current Aspose.Cells version.
            // If size optimization is required, consider using other available options.

            // Save the workbook as PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
