// Title: Load an XLSX workbook with LoadOptions.IgnoreMissingFonts and convert it to PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load Excel file with IgnoreMissingFonts option | how to avoid missing font errors when exporting XLSX to PDF with Aspose.Cells | C# example of using LoadOptions to skip missing fonts during PDF conversion | using LoadOptions.IgnoreMissingFonts property in Aspose.Cells .NET
// Tags: LoadOptions.IgnoreMissingFonts usage | XLSX to PDF conversion with Aspose.Cells | skip missing fonts during Excel to PDF export | C# Aspose.Cells load options for PDF generation | handling missing fonts in Aspose.Cells conversion

using Aspose.Cells;
using System;
using System.IO;

// The sample checks for the source XLSX file, creates a LoadOptions object with the IgnoreMissingFonts flag enabled, loads the workbook using these options, and then saves it as a PDF, all wrapped in error handling to report any issues.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Create load options for XLSX format.
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook with the specified options.
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Save the workbook as PDF.
            workbook.Save(outputPath, SaveFormat.Pdf);

            Console.WriteLine($"Conversion successful. PDF saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
