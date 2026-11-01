// Title: Convert an Excel .xlsx workbook to CSV while preserving column layout and omitting row numbers with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export .xlsx to .csv without row numbers | How to save an Excel workbook as CSV while keeping column order using Aspose.Cells | C# convert Excel file to CSV without adding row headers with Aspose.Cells | Example of SaveFormat.Csv in Aspose.Cells that omits row numbers | Remove row numbers when converting Excel to CSV in .NET with Aspose.Cells
// Tags: Aspose.Cells CSV conversion C# | Workbook.SaveFormat.Csv Aspose.Cells | preserve column order CSV Aspose.Cells | exclude row numbers Excel to CSV | Aspose.Cells .NET CSV export options

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the input .xlsx file, loads it into an Aspose.Cells Workbook, and saves it directly as a .csv file using SaveFormat.Csv. The resulting CSV retains the original column arrangement and does not include row numbers, with error handling for missing files or runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.csv";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook directly as CSV (no additional options required)
            workbook.Save(outputPath, SaveFormat.Csv);
            Console.WriteLine($"Workbook successfully saved as CSV to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
