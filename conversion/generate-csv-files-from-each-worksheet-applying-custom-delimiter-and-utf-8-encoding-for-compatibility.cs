// Title: Export each worksheet of an Excel workbook to separate CSV files with a custom semicolon delimiter and UTF-8 encoding using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export all worksheets to individual CSV files with semicolon separator | How to save each sheet of an .xlsx as UTF-8 CSV using Aspose.Cells .NET | C# generate CSV per worksheet with custom delimiter using TxtSaveOptions | Iterate workbook worksheets and save as CSV with UTF-8 encoding in Aspose.Cells | Set custom CSV separator when converting Excel sheets to CSV in C#
// Tags: export worksheet to CSV custom delimiter Aspose.Cells | TxtSaveOptions CSV separator configuration | UTF-8 CSV output Aspose.Cells | iterate workbook worksheets C# | save active sheet as CSV Aspose.Cells

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The C# program verifies the input Excel file, loads it with Aspose.Cells, and loops through every worksheet. For each sheet it activates the worksheet, configures TxtSaveOptions to use a semicolon as the CSV separator and UTF-8 encoding, then saves the sheet as a separate CSV file named after the worksheet. Errors are caught and reported.
class CsvExport
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Define the custom delimiter you want to use in the CSV files
            char customDelimiter = ';';

            // Iterate through each worksheet in the workbook
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                // Set the current worksheet as the active sheet
                workbook.Worksheets.ActiveSheetIndex = sheetIndex;
                Worksheet sheet = workbook.Worksheets[sheetIndex];

                // Configure CSV save options using TxtSaveOptions (the class for CSV in current API)
                TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
                {
                    // Apply the custom delimiter
                    Separator = customDelimiter,
                    // Use UTF-8 encoding for compatibility
                    Encoding = Encoding.UTF8
                    // Note: Saving a workbook to CSV automatically saves only the active sheet,
                    // so no explicit SaveOnlyActiveSheet property is required.
                };

                // Build the output file name (e.g., "Sheet1.csv")
                string outputFileName = $"{sheet.Name}.csv";

                try
                {
                    // Save the active worksheet as a CSV file with the specified options
                    workbook.Save(outputFileName, csvOptions);
                    Console.WriteLine($"Worksheet \"{sheet.Name}\" saved as \"{outputFileName}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save worksheet \"{sheet.Name}\": {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
