// Title: Export an Excel workbook to CSV while preserving empty cells as empty strings using Aspose.Cells for .NET
// Common Searches: How can I export an Excel file to CSV in C# while keeping blank cells using Aspose.Cells? | Aspose.Cells C# preserve empty cells when saving workbook as CSV | Save Excel workbook to CSV with consistent column count using TxtSaveOptions | C# CSV export from Excel that retains empty strings for empty cells | Configure Aspose.Cells CSV options to avoid dropping empty columns
// Tags: Aspose.Cells CSV export with empty cell handling | TxtSaveOptions SaveFormat.Csv separator setting | C# workbook to CSV conversion maintaining column count | Excel to CSV conversion preserving blank cells Aspose.Cells | Aspose.Cells CSV save options for consistent columns

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example checks for the input Excel file, loads it into an Aspose.Cells Workbook, sets up TxtSaveOptions for CSV with a comma separator, ensures the output directory exists, and saves the workbook as a CSV file while keeping empty cells as empty strings to maintain column alignment.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Define the column separator (comma for CSV)
                Separator = ','
                // ExportBlankCells and ExportAllColumns are not available in this version of Aspose.Cells
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as CSV using the configured options
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
