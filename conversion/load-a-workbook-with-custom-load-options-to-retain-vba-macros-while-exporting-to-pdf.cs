// Title: Convert a macro-enabled .xlsm workbook to PDF while preserving VBA macros using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export macro-enabled Excel file to PDF preserving VBA | How to keep VBA macros when converting .xlsm to PDF with Aspose.Cells | Load .xlsm workbook with macros and save as PDF in .NET | C# Aspose.Cells save workbook as PDF without losing macros | Convert macro-enabled Excel to PDF using Aspose.Cells .NET library
// Tags: Aspose.Cells .xlsm PDF export | macro-enabled workbook loading C# | save workbook as PDF Aspose.Cells | VBA macro preservation .NET | Excel macro-enabled PDF generation

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the .xlsm file exists, loads it with Aspose.Cells (which automatically retains VBA macros), and saves the workbook as a PDF, handling errors and confirming successful conversion.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsm";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook (VBA macros are retained automatically for .xlsm files)
            Workbook workbook = new Workbook(inputPath);

            // Export the workbook to PDF
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Successfully saved PDF to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading or saving issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
