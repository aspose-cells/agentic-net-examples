// Title: Export a single worksheet from an Excel file to a UTF-8 tab-delimited text file using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells export only one worksheet to a tab-separated .txt file | How to save an Excel sheet as UTF-8 text with tab delimiters using Aspose.Cells | Aspose.Cells TxtSaveOptions set separator to tab and encoding to UTF-8 | Export specific sheet from .xlsx to .txt in .NET with Aspose.Cells | Save active worksheet as UTF-8 tab delimited text using Aspose.Cells API
// Tags: Aspose.Cells TxtSaveOptions configuration | worksheet to tab delimited text conversion | UTF-8 encoding for Excel text export | C# export active sheet as .txt | Aspose.Cells save specific sheet as text

using Aspose.Cells;
using System;
using System.IO;
using System.Text;

// The example loads an Excel workbook, selects a worksheet (by name or index), makes it the active sheet, configures TxtSaveOptions with a tab separator and UTF-8 encoding, and saves only that worksheet to a .txt file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.txt";

        try
        {
            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the worksheet to export (by name or index)
            Worksheet worksheet = workbook.Worksheets["Sheet1"]; // or workbook.Worksheets[0];

            // Set the worksheet to be the active sheet so that only it is exported
            workbook.Worksheets.ActiveSheetIndex = worksheet.Index;

            // Configure TXT save options: tab delimiter and UTF‑8 encoding
            TxtSaveOptions txtOptions = new TxtSaveOptions
            {
                Separator = '\t',               // Use tab as the column delimiter
                Encoding = Encoding.UTF8        // Set UTF‑8 encoding
                // SheetIndex is not required; the active sheet will be exported
            };

            // Export the selected worksheet to a TXT file
            workbook.Save(outputPath, txtOptions);
            Console.WriteLine($"Worksheet exported successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
