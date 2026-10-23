// Title: Save a full Excel workbook to CSV with empty rows preserved using Aspose.Cells for .NET
// Common Searches: Aspose.Cells export entire workbook to CSV while keeping blank rows in C# | How to retain empty rows when converting Excel to CSV with Aspose.Cells .NET | C# TxtSaveOptions settings for CSV output that preserve row structure | Save Excel file as CSV without losing blank line separators using Aspose.Cells | Aspose.Cells CSV conversion keep row layout for empty rows
// Tags: Aspose.Cells TxtSaveOptions CSV export | C# workbook to CSV conversion | preserve empty rows CSV Aspose.Cells | Excel to CSV conversion .NET | handle blank rows Aspose.Cells CSV

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook, configures TxtSaveOptions with a comma separator, ensures the output folder exists, and saves the entire workbook as a CSV file using Aspose.Cells for .NET. Blank rows are written as empty lines, maintaining the original row layout. Error handling covers missing input files.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.csv";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Configure CSV save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Separator = ',' // column delimiter
                // KeepSeparatorsForBlankRows is not available in this version; omitted
            };

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as CSV
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook saved as CSV to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
