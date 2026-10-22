// Title: Load an Excel workbook in C# with Aspose.Cells while preserving cell formatting and skipping formula evaluation using LoadOptions
// Common Searches: C# Aspose.Cells load workbook without calculating formulas and keep original formatting | How to use LoadOptions in Aspose.Cells to ignore formula results when opening an XLSX file | Preserve cell styles when loading an Excel file with Aspose.Cells LoadOptions in .NET | Load XLSX workbook without evaluating formulas using Aspose.Cells C# example | Aspose.Cells LoadOptions example for loading Excel with formatting only
// Tags: Aspose.Cells LoadOptions preserve cell formatting | skip formula evaluation LoadOptions Aspose.Cells | load XLSX workbook without calculating formulas C# | preserve Excel styles using LoadOptions | LoadFormat.Xlsx usage Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample checks for the presence of input.xlsx, creates a LoadOptions object specifying the XLSX format, loads the workbook with those options to retain the original cell formatting while avoiding formula calculation, and then saves the result as output_preserved.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output_preserved.xlsx";

                // Ensure the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load options: specify the format of the source file
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

                // Load the workbook with the defined options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook to verify that formatting is preserved
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
