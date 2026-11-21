// Title: How to save a localized Excel workbook as XLSX preserving formatting, comments, and cell styles using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells save workbook as XLSX keep original formatting and comments | preserve cell styles when converting localized Excel file to XLSX with Aspose.Cells | Aspose.Cells .NET load existing workbook and export to XLSX without losing formatting | how to retain comments in Excel after saving with Aspose.Cells in C# | save localized Excel workbook to XLSX preserving all features using Aspose.Cells API
// Tags: Aspose.Cells save workbook as XLSX | preserve Excel formatting Aspose.Cells | retain cell comments Aspose.Cells | localized Excel conversion .NET | cell style preservation Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for an input.xlsx file, loads it into an Aspose.Cells Workbook, and saves it as output.xlsx in XLSX format while preserving all original formatting, comments, and cell styles, with proper exception handling.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the source file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook.
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook in XLSX format, preserving all original features.
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine($"Workbook successfully saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
