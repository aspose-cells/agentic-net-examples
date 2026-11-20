// Title: Export an Excel worksheet to a tab‑delimited text file while preserving numeric formatting with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# export worksheet to tab delimited .txt preserving number format | How to save an Excel sheet as tab‑separated values using TxtSaveOptions in .NET | C# code to write the active worksheet to a .txt file with tab separator and keep numeric formatting | Export Excel data to a tab‑delimited text file using Aspose.Cells without losing formatting | Save workbook as TabDelimited text in C# while maintaining cell number formats
// Tags: Aspose.Cells TabDelimited TxtSaveOptions | C# worksheet to .txt with tab separator | numeric format preservation in text export | active sheet export as .txt using Aspose.Cells | tab‑separated values export .NET

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the source Excel file, loads it into a Workbook, activates the first worksheet, configures TxtSaveOptions for TabDelimited with a tab character, and saves the active sheet to a .txt file while retaining the original numeric formatting.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook containing the data to export
            Workbook workbook = new Workbook(inputPath);

            // Set the first worksheet as the active sheet (if any)
            if (workbook.Worksheets.Count > 0)
            {
                workbook.Worksheets.ActiveSheetIndex = 0;
            }

            // Configure save options for a tab‑delimited text file
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.TabDelimited)
            {
                Separator = '\t' // Use tab as the column delimiter
            };

            // Export the active worksheet to a tab‑delimited text file
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Worksheet exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
