// Title: C# – Convert Semicolon‑Delimited CSV to XLSX with Aspose.Cells TxtLoadOptions
// Description: Demonstrates how to set TxtLoadOptions.Separator to ';', load a CSV file using Aspose.Cells, and save the workbook as an XLSX document in a single step.
// Keywords: Aspose.Cells TxtLoadOptions separator | semicolon CSV to XLSX C# | custom delimiter CSV Aspose | load CSV with TxtLoadOptions | save workbook as XLSX | European CSV conversion | Aspose.Cells CSV import example
// Common Searches: Aspose.Cells set CSV delimiter to semicolon | C# convert CSV with custom separator to Excel | TxtLoadOptions separator usage | How to load semicolon‑separated CSV in .NET | Batch convert CSV files to XLSX with Aspose
// Developer Intent: Configure a custom CSV delimiter, import the file, and export it as an XLSX workbook.
// Use Cases: Transform European‑style CSV reports (semicolon‑separated) into Excel for analysis. | Automate nightly conversion of multiple semicolon‑delimited CSV files to XLSX in a CI pipeline. | Read CSV data with a non‑standard delimiter, apply Aspose.Cells formatting, and generate a shareable Excel file.
// AI Prompts: Generate C# code that uses Aspose.Cells TxtLoadOptions to load a CSV file with a semicolon delimiter and save it as XLSX. | Create a script that scans a directory for *.csv files, converts each using TxtLoadOptions.Separator = ';', and logs conversion results. | Explain how TxtLoadOptions can be configured for different delimiters and how to preserve data types when converting CSV to XLSX with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToXlsx
{
    // Demonstrates how to set TxtLoadOptions.Separator to ';', load a CSV file using Aspose.Cells, and save the workbook as an XLSX document in a single step.
    public class CsvToXlsxConverter
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Verify input file exists
            if (!File.Exists(csvPath))
            {
                Console.WriteLine($"Input file not found: {csvPath}");
                return;
            }

            // Create load options and set the separator to semicolon
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = ';'
            };

            // Load the CSV file using the configured options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Export the loaded workbook to XLSX format
            string xlsxPath = "output.xlsx";
            workbook.Save(xlsxPath, SaveFormat.Xlsx);

            Console.WriteLine($"Conversion completed. Output saved to {xlsxPath}");
        }
    }
}
