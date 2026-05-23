using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

namespace CsvToOdsBatchConversion
{
    class Program
    {
        static void Main()
        {
            // Define CSV files to process
            string[] csvFiles = { "data1.csv", "data2.csv", "data3.csv" };

            // Create sample CSV content for each file (if not already present)
            foreach (string csvPath in csvFiles)
            {
                if (!File.Exists(csvPath))
                {
                    // Simple two‑column sample data
                    string sampleData = "Name,Score\nAlice,85\nBob,92\nCharlie,78";
                    File.WriteAllText(csvPath, sampleData);
                }

                // Determine the output ODS file name
                string odsPath = Path.ChangeExtension(csvPath, ".ods");

                // Load options for CSV format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

                // Save options for ODS with default ODF version
                OdsSaveOptions saveOptions = new OdsSaveOptions();

                // Convert CSV to ODS using the utility method
                ConversionUtility.Convert(csvPath, loadOptions, odsPath, saveOptions);

                // Verify conversion by loading the generated ODS file
                Workbook verificationWorkbook = new Workbook(odsPath);
                Console.WriteLine($"Converted '{csvPath}' to '{odsPath}'. Worksheets count: {verificationWorkbook.Worksheets.Count}");
            }

            // Optional: clean up temporary CSV files
            // foreach (string csvPath in csvFiles) { if (File.Exists(csvPath)) File.Delete(csvPath); }
        }
    }
}