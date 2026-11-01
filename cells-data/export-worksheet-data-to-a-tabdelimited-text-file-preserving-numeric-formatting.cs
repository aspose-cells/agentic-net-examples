// Title: Export active worksheet to a tab‑delimited .txt file while preserving numeric formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an .xlsx workbook with Aspose.Cells, sets the first worksheet as the active sheet, and saves it as a tab‑delimited .txt file preserving the original number formats. | Show how to configure Aspose.Cells TxtSaveOptions for TabDelimited output with a tab character as the separator and export only the active worksheet in C#. | Add comprehensive error handling to a C# program that exports a worksheet to a tab‑delimited text file using Aspose.Cells, including file‑existence verification and detailed exception logging.
// Common Searches: C# Aspose.Cells export only active sheet to tab delimited text preserving number format | How to save an Excel worksheet as a .txt file with tabs using Aspose.Cells .NET | Aspose.Cells TxtSaveOptions TabDelimited numeric formatting example in C# | Export Excel data to a tab‑separated values file with Aspose.Cells C#
// Tags: Aspose.Cells active sheet TabDelimited export | TxtSaveOptions tab delimiter C# | numeric format preservation Aspose.Cells txt save | C# load workbook with Aspose.Cells | export worksheet to tab‑separated .txt Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace ExportWorksheetToTabDelimited
{
    // The example loads an Excel workbook (input.xlsx) with Aspose.Cells, sets the first worksheet as the active sheet, configures TxtSaveOptions for TabDelimited format using a tab separator, and saves the active worksheet to output.txt while preserving numeric formatting. It also includes checks for missing input files and generic exception handling.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source Excel file
            string inputFile = "input.xlsx";

            // Path for the exported tab‑delimited text file
            string outputFile = "output.txt";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook from the Excel file
                Workbook workbook = new Workbook(inputFile);

                // Set the first worksheet as the active sheet (only active sheet is saved in TabDelimited format)
                workbook.Worksheets.ActiveSheetIndex = 0;

                // Configure TXT save options for tab‑delimited output
                TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.TabDelimited)
                {
                    // Use tab as the column separator
                    Separator = '\t'
                };

                // Save the active worksheet as a tab‑delimited text file
                workbook.Save(outputFile, saveOptions);
                Console.WriteLine($"Worksheet exported successfully to {outputFile}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
