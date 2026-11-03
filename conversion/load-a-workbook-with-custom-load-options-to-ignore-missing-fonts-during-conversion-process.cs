// Title: How to load an Excel workbook with LoadOptions that ignore missing fonts and convert it to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load workbook with missing font fallback and save as PDF | ignore unavailable fonts when converting XLSX to PDF with Aspose.Cells .NET | custom LoadOptions to suppress font errors during Excel to PDF conversion in C#
// Tags: load workbook with missing font fallback Aspose.Cells | Excel to PDF conversion ignore font errors C# | custom LoadOptions for font handling Aspose.Cells | suppress missing fonts during workbook load .NET | Aspose.Cells PDF export with font fallback

using Aspose.Cells;
using System;
using System.IO;

// This example shows how to create LoadOptions that ignore missing fonts, load an XLSX workbook with those options, and then save the workbook as a PDF, handling any font‑related issues during the conversion.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook with default options
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook to PDF
            string outputPath = "output.pdf";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
