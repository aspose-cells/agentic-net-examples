// Title: Export an Excel worksheet to a UTF-8 CSV file as plain text while preserving cell data types using Aspose.Cells for .NET
// Common Searches: how to save an Excel worksheet as CSV with text format using Aspose.Cells C# | preserve numeric and date values as text when exporting to CSV with Aspose.Cells | Aspose.Cells C# export active worksheet to UTF-8 CSV file | set TxtSaveOptions for CSV export to keep cell data types in Aspose.Cells | convert XLSX to CSV without losing cell formatting using Aspose.Cells .NET
// Tags: Aspose.Cells TxtSaveOptions CSV example | CSV export preserving cell data types Aspose.Cells | UTF8 encoding for CSV output Aspose.Cells | active worksheet CSV save Aspose.Cells | C# Aspose.Cells SaveFormat.Csv usage

using System;
using System.IO;
using System.Text;
using Aspose.Cells;

// The code verifies the input XLSX file, loads it into an Aspose.Cells Workbook, activates the first worksheet, configures TxtSaveOptions for CSV format with UTF-8 encoding, and saves the active worksheet as a CSV file while retaining the original cell data types as text.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Ensure the first worksheet is the active one (CSV export uses the active sheet)
            workbook.Worksheets.ActiveSheetIndex = 0;

            // Configure CSV (text) save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8   // set desired encoding
                // Additional options can be set here if needed
            };

            // Export the selected worksheet to a CSV file
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Worksheet exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
