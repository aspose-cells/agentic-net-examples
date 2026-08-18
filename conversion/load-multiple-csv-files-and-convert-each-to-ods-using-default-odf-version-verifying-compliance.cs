// Title: Batch convert CSV files to ODS with Aspose.Cells for .NET (default ODF version)
// Description: Scans a folder for *.csv files, loads each into an Aspose.Cells Workbook using LoadOptions, saves it as an ODS document with OdsSaveOptions (default ODF version), and confirms the output file exists while handling per‑file errors.
// Keywords: Aspose.Cells | CSV to ODS conversion | batch conversion C# | LoadOptions CSV | OdsSaveOptions | default ODF version | file existence verification | .NET spreadsheet automation | convert multiple CSV files
// Common Searches: Aspose.Cells batch CSV to ODS example | C# convert folder of CSV files to ODS | verify ODS files after conversion Aspose | default ODF version when saving ODS | load CSV with Aspose.Cells LoadOptions
// Developer Intent: Automatically transform every CSV file in a directory into an ODS spreadsheet using Aspose.Cells and ensure each result is created successfully.
// Use Cases: Nightly job that turns exported CSV reports into LibreOffice‑compatible ODS files. | Migrating legacy CSV datasets to OpenDocument spreadsheets for downstream processing. | CI/CD validation step that checks ODS output files exist after a bulk conversion.
// AI Prompts: Write C# code that iterates over all CSV files in a given folder, loads each with Aspose.Cells LoadOptions, saves them as ODS using the default ODF version, and logs success or failure. | Add functionality to the batch converter that records source CSV path, destination ODS path, and conversion status into a summary CSV report. | Explain how to set a specific ODF version in OdsSaveOptions and how to programmatically validate ODS compliance after saving.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Ods;

namespace CsvToOdsBatchConversion
{
    // Scans a folder for *.csv files, loads each into an Aspose.Cells Workbook using LoadOptions, saves it as an ODS document with OdsSaveOptions (default ODF version), and confirms the output file exists while handling per‑file errors.
    class Program
    {
        static void Main()
        {
            // Folder containing CSV files – adjust as needed
            string inputFolder = @"C:\CsvFiles";

            // Verify the folder exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder does not exist: {inputFolder}");
                return;
            }

            // Get all CSV files in the folder
            string[] csvFiles = Directory.GetFiles(inputFolder, "*.csv");

            if (csvFiles.Length == 0)
            {
                Console.WriteLine("No CSV files found to convert.");
                return;
            }

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // ---------- Load CSV ----------
                    // Create load options specifying CSV format
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                    // Load the CSV file into a workbook using the load options
                    Workbook workbook = new Workbook(csvPath, loadOptions);

                    // ---------- Save as ODS ----------
                    // Create ODS save options – default ODF version will be used
                    OdsSaveOptions saveOptions = new OdsSaveOptions();

                    // Determine output ODS file path (same name, .ods extension)
                    string odsPath = Path.ChangeExtension(csvPath, ".ods");

                    // Save the workbook as ODS with the specified options
                    workbook.Save(odsPath, saveOptions);

                    // ---------- Verify conversion ----------
                    // Simple verification: check that the ODS file now exists
                    if (File.Exists(odsPath))
                    {
                        Console.WriteLine($"Successfully converted '{Path.GetFileName(csvPath)}' to ODS.");
                    }
                    else
                    {
                        Console.WriteLine($"Failed to create ODS file for '{Path.GetFileName(csvPath)}'.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing '{Path.GetFileName(csvPath)}': {ex.Message}");
                }
            }
        }
    }
}
