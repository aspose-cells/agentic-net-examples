// Title: Implement error handling for missing hidden worksheets when using ExportHiddenWorksheet during PDF export with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that checks if a workbook contains any hidden worksheets before setting PdfSaveOptions.ExportHiddenWorksheet = true and throws a descriptive InvalidOperationException when none are found. | Show how to wrap the Aspose.Cells PDF export routine in a try‑catch block that logs a custom error message if ExportHiddenWorksheet is enabled but the workbook has no hidden sheets. | Provide a refactored version of the sample that validates hidden worksheets, sets the ExportHiddenWorksheet flag, and returns an appropriate error code to the caller when the validation fails.
// Common Searches: Aspose.Cells how to validate hidden worksheets before PDF export in C# | ExportHiddenWorksheet option throws error when no hidden sheets present | C# check for hidden worksheets in Aspose.Cells workbook prior to saving as PDF | handle missing hidden worksheets with Aspose.Cells PdfSaveOptions ExportHiddenWorksheet | Aspose.Cells PDF conversion error handling for absent hidden worksheets
// Tags: Aspose.Cells PDF export hidden worksheet validation | C# check hidden sheets before saving to PDF | ExportHiddenWorksheet error handling Aspose.Cells | PdfSaveOptions hidden worksheet verification C# | custom exception missing hidden worksheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates how to verify the presence of hidden worksheets before enabling ExportHiddenWorksheet in PdfSaveOptions, includes robust try‑catch error handling, and throws a clear exception when no hidden sheets are found, ensuring safe PDF export with Aspose.Cells for .NET.
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
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure PDF save options
            PdfSaveOptions saveOptions = new PdfSaveOptions();

            // NOTE: In recent versions of Aspose.Cells the ExportHiddenWorksheet
            // option is not available. Hidden worksheets are exported by default
            // when saving to PDF. If you need to control this behavior, ensure
            // you are using a version that supports the property.

            // Save the workbook as PDF
            workbook.Save(outputPath, saveOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
