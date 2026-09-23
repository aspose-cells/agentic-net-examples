// Title: Import a CSV file with a custom delimiter and header row into an Aspose.Cells workbook and save as XLSX using C#
// AI Prompts: Read a CSV document in C# with Aspose.Cells TxtLoadOptions, specifying a custom separator and treating the first line as column headers. | Generate an XLSX workbook from a delimited CSV source in .NET by applying TxtLoadOptions to preserve the header row.
// Common Searches: C# Aspose.Cells load csv file with semicolon delimiter and first row as column names | How to import a CSV with custom separator into a workbook using Aspose.Cells | Aspose.Cells TxtLoadOptions example for CSV with header row | Convert CSV to XLSX in .NET preserving header using Aspose.Cells | Load CSV into Aspose.Cells workbook specifying delimiter and header option
// Tags: TxtLoadOptions custom delimiter CSV import | Aspose.Cells CSV import preserving headers | Create Workbook from delimited text Aspose | Save workbook as XLSX using Aspose.Cells | C# import CSV into Aspose.Cells workbook

using System;
using System.IO;
using Aspose.Cells;

// The code checks for the CSV file, configures TxtLoadOptions with a custom separator, loads the CSV into an Aspose.Cells Workbook while treating the first row as headers, and then saves the workbook as an XLSX file, handling any errors that may occur.
class Program
{
    static void Main()
    {
        // Path to the CSV file
        string csvPath = "data.csv";

        // Verify that the CSV file exists
        if (!File.Exists(csvPath))
        {
            Console.WriteLine($"Error: File '{csvPath}' not found.");
            return;
        }

        // Specify the delimiter character (e.g., comma)
        char delimiter = ',';

        try
        {
            // Configure CSV load options using TxtLoadOptions (supports custom delimiter)
            TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.Csv)
            {
                Separator = delimiter // Set custom delimiter
            };

            // Load the CSV file into a workbook using the defined options
            Workbook workbook = new Workbook(csvPath, loadOptions);

            // Save the workbook to another format, e.g., XLSX
            string outputPath = "output.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
