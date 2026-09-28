// Title: Batch convert CSV files with a custom delimiter to XLSX using Aspose.Cells for .NET
// AI Prompts: Create a C# console application that scans a directory for *.csv files, loads each file with a semicolon delimiter via Aspose.Cells TxtLoadOptions, and saves the result as an .xlsx workbook. | Extend the sample to accept the delimiter character and the output folder as command‑line parameters, then perform the CSV‑to‑XLSX conversion for all files in the source folder.
// Common Searches: asp.net convert all csv files in a folder to xlsx with custom delimiter | c# load semicolon delimited csv using Aspose.Cells TxtLoadOptions | automate csv to excel conversion for multiple files in .NET | Aspose.Cells TxtLoadOptions separator example | convert directory of csv to xlsx programmatically
// Tags: csv to xlsx batch processing Aspose.Cells | TxtLoadOptions separator property usage | multiple CSV files conversion C# | save workbook as Xlsx format .NET | semicolondelimited CSV import Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToXlsxBatch
{
    // The program enumerates every .csv file in a specified folder, loads each using Aspose.Cells TxtLoadOptions with a custom semicolon separator, and saves the workbook as an .xlsx file with the same base name.
    class Program
    {
        static void Main()
        {
            // Directory containing the CSV files
            string inputDirectory = @"C:\Data\CsvFiles";

            // Get all CSV files in the directory
            string[] csvFiles = Directory.GetFiles(inputDirectory, "*.csv");

            foreach (string csvPath in csvFiles)
            {
                // Create load options for CSV with a custom delimiter (e.g., semicolon)
                TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv);
                loadOptions.Separator = ';';               // custom delimiter
                loadOptions.ConvertNumericData = true;     // optional: convert numbers

                // Load the CSV file into a workbook using the specified options
                Workbook workbook = new Workbook(csvPath, loadOptions);

                // Determine the output XLSX file path (same name, different extension)
                string xlsxPath = Path.ChangeExtension(csvPath, ".xlsx");

                // Save the workbook as XLSX
                workbook.Save(xlsxPath, SaveFormat.Xlsx);
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
