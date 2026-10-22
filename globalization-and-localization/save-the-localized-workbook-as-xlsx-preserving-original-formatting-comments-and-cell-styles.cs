// Title: How to save a localized Excel workbook to XLSX while preserving formatting, comments, and cell styles using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# save workbook to XLSX without losing cell formatting and comments | preserve original styles when exporting a localized Excel file to XLSX using .NET | OoxmlSaveOptions example for keeping comments and formatting in C# | export a localized workbook to XLSX while retaining all cell styles with Aspose.Cells | C# code to save Excel file as XLSX preserving formatting and comments
// Tags: OoxmlSaveOptions retain cell styles | localized workbook export to XLSX | C# save Excel with comments Aspose.Cells | XLSX save preserving visual layout | Aspose.Cells keep formatting on save

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, configures OoxmlSaveOptions for the XLSX format, and saves the workbook while retaining the original formatting, comments, and cell styles.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure save options for XLSX (OOXML) while preserving formatting and styles
            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions(SaveFormat.Xlsx);

            // Save the workbook with the specified options
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
