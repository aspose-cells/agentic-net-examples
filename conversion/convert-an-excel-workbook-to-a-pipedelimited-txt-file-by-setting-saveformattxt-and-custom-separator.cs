// Title: How to export an Excel workbook to a pipe‑delimited TXT file with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells export Excel to pipe separated text file | set custom column separator when saving workbook as txt using Aspose.Cells | convert .xlsx to pipe delimited .txt with Aspose.Cells .NET | TxtSaveOptions separator property example in C# | save workbook as pipe delimited txt using Aspose.Cells API
// Tags: Aspose.Cells TxtSaveOptions pipe separator | C# export Excel to custom delimited TXT | save workbook as pipe delimited text .NET | Excel to pipe separated file Aspose.Cells | custom delimiter txt export Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source XLSX file, loads it into an Aspose.Cells Workbook, configures TxtSaveOptions with a pipe ('|') as the column separator, and saves the workbook as a pipe‑delimited TXT file, handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.txt";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the Excel workbook from the specified file
            Workbook workbook = new Workbook(inputPath);

            // Configure TXT save options to use a pipe (|) as the column separator
            TxtSaveOptions txtOptions = new TxtSaveOptions
            {
                Separator = '|'
            };

            // Save the workbook as a pipe‑delimited TXT file
            workbook.Save(outputPath, txtOptions);
            Console.WriteLine($"Workbook successfully saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
