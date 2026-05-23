using System;
using System.IO;
using Aspose.Cells;

class AutoPopulateExample
{
    static void Main()
    {
        // Path to the source CSV (or any delimited text) that may exceed a sheet's row limit
        string inputFile = "largeData.csv";

        // Path where the resulting workbook will be saved
        string outputFile = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputFile))
            {
                Console.WriteLine($"Input file not found: {inputFile}");
                return;
            }

            // Configure loading options to automatically continue data on a new worksheet
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                // When rows exceed the maximum per sheet, create/append a new sheet
                ExtendToNextSheet = true
            };

            // Load the CSV into a workbook using the configured options
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Auto‑fit columns on each generated worksheet for better readability
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.AutoFitColumns();
            }

            // Save the workbook with all data distributed across multiple sheets as needed
            workbook.Save(outputFile);
            Console.WriteLine($"Workbook saved successfully to {outputFile}");
        }
        catch (Exception ex)
        {
            // Catch any runtime errors and display a concise message
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}