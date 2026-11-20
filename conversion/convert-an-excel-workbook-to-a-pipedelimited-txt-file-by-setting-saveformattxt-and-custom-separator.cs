// Title: How to export an Excel workbook to a pipe‑delimited TXT file using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# save workbook as pipe delimited text file | C# TxtSaveOptions set custom separator to pipe for Excel to TXT conversion | Convert .xlsx to pipe‑separated .txt using Aspose.Cells API | How to specify delimiter when exporting Excel to TXT with Aspose.Cells in .NET | Example code for saving Excel as pipe delimited text using Aspose.Cells
// Tags: Aspose.Cells TxtSaveOptions custom delimiter | pipe delimited TXT export from Excel C# | Excel workbook to text file conversion Aspose | C# save workbook as text with custom separator | Aspose.Cells export to pipe separated values

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // Loads an Excel workbook, configures TxtSaveOptions with a pipe ('|') separator, and saves the workbook as a pipe‑delimited .txt file using Aspose.Cells for .NET, including file existence checks and exception handling.
    class ExcelToPipeDelimitedTxt
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.txt";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the source Excel workbook
                Workbook workbook = new Workbook(inputPath);

                // Configure TXT save options with pipe (|) as the delimiter
                TxtSaveOptions saveOptions = new TxtSaveOptions();
                saveOptions.Separator = '|';

                // Save the workbook as a pipe‑delimited TXT file
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved as pipe‑delimited text: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
