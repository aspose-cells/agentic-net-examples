// Title: Load an .xlsx workbook and export it to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells code to read an .xlsx file and save as PDF | How to use Aspose.Cells SaveFormat.Pdf in a .NET console app | Check file existence before converting Excel to PDF with Aspose.Cells | Example of exception handling during Excel to PDF conversion in C# | Aspose.Cells workbook to PDF conversion tutorial for .NET developers
// Tags: Aspose.Cells load workbook from .xlsx | Aspose.Cells export workbook to PDF | C# file existence verification for Excel conversion | Aspose.Cells SaveFormat.Pdf parameter usage | Aspose.Cells error handling during PDF export

using System;
using System.IO;
using Aspose.Cells;

// The program verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, and saves the workbook as output.pdf using the SaveFormat.Pdf option, with try‑catch handling to report any runtime errors.
class Program
{
    static void Main()
    {
        // Input Excel file and output PDF file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the existing Excel file
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook to PDF format
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
