// Title: Convert an Excel workbook to PDF and set the author metadata using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set workbook author before exporting to PDF | How to embed creator name when converting Excel to PDF with Aspose.Cells | Aspose.Cells PDF export preserving document metadata in .NET | Example of using workbook.Settings.Author for PDF conversion in C# | Convert Excel file to PDF and add author property using Aspose.Cells library
// Tags: Aspose.Cells workbook to PDF conversion | workbook.Settings.Author usage | PDF export with author metadata Aspose.Cells | C# Excel to PDF conversion preserving metadata | Aspose.Cells document properties for PDF output

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an Excel workbook, assigns the author via the Settings.Author property, ensures the output folder exists, and saves the workbook as a PDF using SaveFormat.Pdf, handling missing files and runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Set the author property of the workbook (use Settings property)
            workbook.Settings.Author = "John Doe";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Convert and save the workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook successfully saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
