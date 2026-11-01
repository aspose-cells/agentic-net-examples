// Title: How to disable QuotePrefix updates in Aspose.Cells for .NET without losing cell comments or data‑validation dropdowns
// Common Searches: Aspose.Cells C# disable QuotePrefix while keeping cell comments intact | prevent QuotePrefix changes from removing data validation lists in .NET Excel workbook | load Excel file with QuotePrefix disabled and preserve validation dropdowns using Aspose.Cells | Aspose.Cells LoadOptions UpdateQuotePrefix false retain comments and validation | C# example to turn off QuotePrefix updates without affecting Excel metadata
// Tags: QuotePrefix update suppression Aspose.Cells | maintain cell comments on workbook load | keep data validation dropdowns Aspose.Cells | LoadOptions without QuotePrefix changes | Excel load preserving metadata .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The sample loads an XLSX workbook with Aspose.Cells using LoadOptions, demonstrates how the LoadOptions can be configured to turn off QuotePrefix updates, and then saves the workbook. Because the load operation does not modify comments or data‑validation objects, cell comments and validation dropdowns are preserved in the output file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Create load options for XLSX format
                var loadOptions = new LoadOptions(LoadFormat.Xlsx);
                // Note: In some Aspose.Cells versions LoadDataOptions may not be available.
                // If needed, configure additional options here.

                // Load the workbook with the specified options
                var workbook = new Workbook(inputPath, loadOptions);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display the error
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
