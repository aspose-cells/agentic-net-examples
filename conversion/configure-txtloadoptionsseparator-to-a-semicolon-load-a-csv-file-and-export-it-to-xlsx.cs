// Title: C# – Convert a semicolon‑delimited CSV to XLSX using Aspose.Cells TxtLoadOptions
// Description: Demonstrates how to configure TxtLoadOptions.Separator to ';', load a CSV file with a custom delimiter into an Aspose.Cells Workbook, and save the result as an XLSX workbook. Includes file‑existence checking and basic exception handling for robust .NET applications.
// Keywords: Aspose.Cells | C# CSV to XLSX | semicolon delimiter | TxtLoadOptions Separator | .NET Excel export | custom CSV delimiter | load CSV Aspose | Excel conversion C# | batch CSV to XLSX
// Common Searches: Aspose.Cells set TxtLoadOptions separator to semicolon | C# convert CSV with ; delimiter to Excel | load CSV with custom delimiter using Aspose.Cells | save loaded CSV workbook as XLSX in .NET | how to handle missing CSV file Aspose.Cells
// Developer Intent: Load a CSV that uses ';' as the field separator and export it to an XLSX file with Aspose.Cells.
// Use Cases: Transform European‑style CSV reports (semicolon separated) into Excel for business analysis. | Automate nightly batch conversion of multiple semicolon‑delimited CSV files to XLSX in a Windows service. | Integrate custom‑delimited CSV ingestion into a data‑processing pipeline that outputs Excel workbooks.
// AI Prompts: Generate C# code that reads a pipe‑delimited CSV with Aspose.Cells TxtLoadOptions and saves it as XLSX. | Explain how to specify character encoding when loading a CSV with a custom separator in Aspose.Cells. | Provide a pattern for logging and retry logic during bulk CSV‑to‑XLSX conversion in a console app.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to configure TxtLoadOptions.Separator to ';', load a CSV file with a custom delimiter into an Aspose.Cells Workbook, and save the result as an XLSX workbook. Includes file‑existence checking and basic exception handling for robust .NET applications.
    public class CsvToXlsxConverter
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }

        public static void Run()
        {
            // Path to the source CSV file (replace with your actual file path)
            string csvPath = "input.csv";

            // Verify that the CSV file exists to avoid FileNotFoundException
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Error: The file '{csvPath}' was not found.");
                return;
            }

            try
            {
                // Create load options and set the separator to a semicolon
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Separator = ';'
                };

                // Load the CSV file into a workbook using the configured options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Save the workbook as an XLSX file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Conversion successful. XLSX saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Handle any runtime exceptions gracefully
                Console.WriteLine($"An error occurred during conversion: {ex.Message}");
            }
        }
    }
}
