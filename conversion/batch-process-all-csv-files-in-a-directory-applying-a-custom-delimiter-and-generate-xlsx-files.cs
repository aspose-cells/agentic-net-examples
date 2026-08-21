// Title: Batch CSV‑to‑XLSX conversion with custom delimiter using Aspose.Cells for .NET
// Description: A C# utility that scans a folder for *.csv files, imports each using a user‑defined delimiter (e.g., ';'), converts numeric values, and saves the result as an .xlsx workbook with the same name. Includes directory validation, error handling, and automatic output folder creation.
// Keywords: Aspose.Cells CSV import | C# batch CSV to XLSX | custom delimiter Excel conversion | process multiple CSV files .NET | folder‑wide CSV to Excel | Aspose.Cells SaveFormat.Xlsx | CSV to Excel automation
// Common Searches: Aspose.Cells batch convert CSV to XLSX | C# import CSV with semicolon delimiter Aspose | convert all CSV files in a directory to Excel | how to use ImportCSV with custom delimiter | automate CSV to Excel conversion .NET
// Developer Intent: Automatically convert every CSV file in a given directory to an XLSX workbook using a specified delimiter.
// Use Cases: Nightly transformation of semicolon‑delimited export reports into Excel for analysts. | ETL step that turns a feed of CSV data files into Excel workbooks for downstream processing. | Desktop utility for users to bulk‑convert data exports from various systems into .xlsx format.
// AI Prompts: Write C# code that uses Aspose.Cells to read all CSV files from a folder, apply a pipe (|) delimiter, and save each as an XLSX file. | Suggest robust logging and exception handling for a batch CSV‑to‑XLSX conversion tool built with Aspose.Cells. | Explain how to preserve original text qualifiers and cell formatting when converting CSV files to Excel with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

// A C# utility that scans a folder for *.csv files, imports each using a user‑defined delimiter (e.g., ';'), converts numeric values, and saves the result as an .xlsx workbook with the same name. Includes directory validation, error handling, and automatic output folder creation.
class CsvToXlsxBatch
{
    static void Main()
    {
        // Directory containing source CSV files
        string inputDirectory = @"C:\InputCsv";

        // Directory where converted XLSX files will be saved
        string outputDirectory = @"C:\OutputXlsx";

        // Custom delimiter to be used when importing CSV files
        string customDelimiter = ";";

        try
        {
            // Verify input directory exists
            if (!Directory.Exists(inputDirectory))
            {
                Console.WriteLine($"Input directory does not exist: {inputDirectory}");
                return;
            }

            // Ensure the output directory exists
            Directory.CreateDirectory(outputDirectory);

            // Process each CSV file in the input directory
            foreach (string csvPath in Directory.GetFiles(inputDirectory, "*.csv"))
            {
                // Verify the CSV file still exists
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"File not found (skipped): {csvPath}");
                    continue;
                }

                try
                {
                    // Create a new empty workbook
                    Workbook workbook = new Workbook();

                    // Get the first worksheet and its cells collection
                    Worksheet worksheet = workbook.Worksheets[0];
                    Cells cells = worksheet.Cells;

                    // Import the CSV using the custom delimiter, convert numeric data, start at cell A1 (row 0, column 0)
                    cells.ImportCSV(csvPath, customDelimiter, true, 0, 0);

                    // Build the output XLSX file path (same name as CSV, different extension)
                    string outputFileName = Path.GetFileNameWithoutExtension(csvPath) + ".xlsx";
                    string xlsxPath = Path.Combine(outputDirectory, outputFileName);

                    // Save the workbook as XLSX
                    workbook.Save(xlsxPath, SaveFormat.Xlsx);

                    Console.WriteLine($"Converted: {csvPath} -> {xlsxPath}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{csvPath}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
