// Title: Convert an Excel workbook to CSV without exporting cell hyperlinks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells and saves it as a CSV while ensuring hyperlinks are not included. | Show how to check for the source file, handle exceptions, and export to CSV with hyperlink suppression using Aspose.Cells. | Demonstrate converting an Excel workbook to CSV in C# and verify that hyperlink data is omitted in the resulting file.
// Common Searches: Aspose.Cells C# save workbook as CSV without hyperlinks | How to ignore hyperlinks when exporting Excel to CSV using Aspose.Cells | C# convert .xlsx to .csv exclude hyperlink data | SaveFormat.Csv does not export hyperlinks Aspose.Cells example | Prevent hyperlink URLs from appearing in CSV output from Aspose.Cells
// Tags: Aspose.Cells CSV export without hyperlinks | C# workbook to CSV conversion | hyperlink exclusion in CSV output | SaveFormat.Csv usage Aspose.Cells | Excel to CSV conversion error handling

using Aspose.Cells;
using System;
using System.IO;

// The program verifies that the input Excel file exists, loads it with Aspose.Cells, and saves it as a CSV using SaveFormat.Csv, which by default omits cell hyperlinks. It includes try‑catch error handling to report any runtime issues.
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

            // Load the source workbook
            Workbook workbook = new Workbook(inputPath);

            // Save the workbook as CSV (hyperlinks are not exported by default)
            workbook.Save(outputPath, SaveFormat.Csv);

            Console.WriteLine($"Workbook successfully saved as CSV to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
