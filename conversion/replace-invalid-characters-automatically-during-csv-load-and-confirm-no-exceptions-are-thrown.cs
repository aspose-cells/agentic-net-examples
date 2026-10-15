// Title: Load a CSV file into an Aspose.Cells workbook in C# while handling invalid characters and preventing exceptions
// Common Searches: how to load a CSV with Aspose.Cells in C# without throwing exceptions | Aspose.Cells C# load CSV and ignore invalid characters during import | C# example for loading CSV into workbook using LoadOptions and catching errors | convert CSV to XLSX with Aspose.Cells while handling malformed data | Aspose.Cells load CSV safely and verify successful workbook creation
// Tags: Aspose.Cells CSV load with exception handling | C# workbook load CSV using LoadOptions | convert CSV to XLSX Aspose.Cells | handle invalid characters Aspose.Cells CSV import | Aspose.Cells robust CSV import | Aspose.Cells data sanitization during CSV load

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the CSV file, loads it into an Aspose.Cells Workbook using LoadOptions for CSV format, catches any exceptions that may arise from invalid characters, and saves the workbook as an XLSX file to confirm successful processing.
class CsvLoadExample
{
    static void Main()
    {
        // Path to the CSV file to be loaded
        string csvPath = "input.csv";

        // Path where the workbook will be saved after loading
        string outputPath = "output.xlsx";

        // Verify that the input CSV file exists
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Input file not found: {csvPath}");
            return;
        }

        try
        {
            // Configure load options for CSV (default options)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Csv);

            // Load the CSV file into a workbook using the configured options
            Workbook workbook = new Workbook(csvPath, loadOptions);
            Console.WriteLine("CSV loaded successfully.");

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to verify that loading succeeded
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error processing CSV: {ex.Message}");
        }
    }
}
