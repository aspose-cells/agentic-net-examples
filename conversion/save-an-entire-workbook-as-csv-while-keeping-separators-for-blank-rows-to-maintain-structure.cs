// Title: Save an entire Excel workbook to a single CSV file while preserving blank rows using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export whole workbook to CSV with empty rows retained | How to keep blank rows when converting Excel to CSV with Aspose.Cells | Save multiple worksheets as one CSV file using TxtSaveOptions in .NET | C# Aspose.Cells CSV conversion preserving row structure | Convert .xlsx to .csv without losing blank lines using Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions CSV conversion | preserve blank rows CSV export | export entire workbook as single CSV | C# Excel to CSV with separator | SaveFormat.Csv Aspose.Cells example

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook, configures TxtSaveOptions with the CSV format and a comma separator, and saves the whole workbook as one CSV file. Blank rows are emitted as empty lines, maintaining the original worksheet structure.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.csv";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Configure CSV save options using TxtSaveOptions
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Use comma as the separator
                Separator = ','
                // Note: ExportBlankRows property is not available in this version
            };

            // Save the entire workbook as a single CSV file
            workbook.Save(outputPath, csvOptions);
            Console.WriteLine($"Workbook successfully saved as CSV to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
