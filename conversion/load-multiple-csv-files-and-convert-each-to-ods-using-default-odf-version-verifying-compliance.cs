// Title: Batch convert all CSV files in a folder to ODS using Aspose.Cells for .NET with default ODF version and verify each output
// AI Prompts: Write C# code that scans a directory, loads each .csv with LoadOptions(LoadFormat.Csv), converts it to .ods using ConversionUtility.Convert and OdsSaveOptions with default settings, and saves the result in an output folder. | Add logic to reload each generated .ods file with OdsLoadOptions, then output the number of worksheets and the ODF strict version used to confirm compliance. | Enhance the program with robust error handling and console logging for missing input folder, empty folder, and per‑file conversion failures.
// Common Searches: convert a folder of CSV documents to ODS spreadsheets using Aspose.Cells C# | Aspose.Cells example for loading CSV and saving as OpenDocument Spreadsheet | how to verify that an ODS file loads correctly after export with Aspose.Cells | default ODF version applied by OdsSaveOptions in Aspose.Cells .NET
// Tags: Aspose.Cells CSV to ODS conversion utility | default ODF version OdsSaveOptions | LoadOptions for CSV files C# | ODS workbook validation after export | error handling for directory based file conversion C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Ods;

namespace CsvToOdsBatchConversion
{
    // The sample scans an input directory for .csv files, converts each to an .ods file using Aspose.Cells' ConversionUtility with default OdsSaveOptions, reloads the generated ODS to ensure it opens, and logs the worksheet count and ODF version. It includes folder existence checks and per‑file error handling.
    class Program
    {
        static void Main()
        {
            // Folder that contains the CSV files to be converted
            string inputFolder = "InputCsv";
            // Folder where the resulting ODS files will be saved
            string outputFolder = "OutputOds";

            // Verify input directory exists
            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine($"Input folder '{inputFolder}' does not exist.");
                return;
            }

            // Ensure output directory exists
            if (!Directory.Exists(outputFolder))
                Directory.CreateDirectory(outputFolder);

            // Get all CSV files in the input folder
            string[] csvFiles = Directory.GetFiles(inputFolder, "*.csv");

            if (csvFiles.Length == 0)
            {
                Console.WriteLine("No CSV files found in the input folder.");
                return;
            }

            foreach (string csvPath in csvFiles)
            {
                try
                {
                    // Determine the ODS file name (same base name, .ods extension)
                    string odsPath = Path.Combine(outputFolder, Path.GetFileNameWithoutExtension(csvPath) + ".ods");

                    // LoadOptions for CSV files (default options are sufficient)
                    LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                    // OdsSaveOptions with default ODF version (None = default, non‑strict)
                    OdsSaveOptions saveOptions = new OdsSaveOptions();

                    // Convert CSV to ODS using the utility method that respects the provided options
                    ConversionUtility.Convert(csvPath, loadOptions, odsPath, saveOptions);

                    // Verify compliance by loading the generated ODS file
                    OdsLoadOptions odsLoadOptions = new OdsLoadOptions();
                    Workbook odsWorkbook = new Workbook(odsPath, odsLoadOptions);

                    // Output verification details
                    Console.WriteLine($"Converted '{Path.GetFileName(csvPath)}' to '{Path.GetFileName(odsPath)}'.");
                    Console.WriteLine($"  Worksheets loaded: {odsWorkbook.Worksheets.Count}");
                    Console.WriteLine($"  ODF version used: {saveOptions.OdfStrictVersion}");
                    Console.WriteLine();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{Path.GetFileName(csvPath)}': {ex.Message}");
                }
            }

            Console.WriteLine("Batch conversion completed.");
        }
    }
}
