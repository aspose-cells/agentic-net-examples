// Title: Load a pipe‑delimited text file using TxtLoadOptions.Separator and convert it to XLSX with Aspose.Cells for .NET (C#)
// AI Prompts: Set TxtLoadOptions.Separator='|' to read a pipe‑separated .txt file into an Aspose.Cells Workbook, then save the workbook as an .xlsx file. | Configure a custom delimiter in TxtLoadOptions, load the source text into a Workbook, and export the result to XLSX format using Aspose.Cells in C#.
// Common Searches: how to set TxtLoadOptions.Separator to pipe character in Aspose.Cells C# | convert pipe delimited txt file to xlsx with Aspose.Cells .NET | load custom delimited text into Aspose.Cells workbook example | Aspose.Cells C# load text file with custom separator and save as Excel | reading pipe separated values using Aspose.Cells TxtLoadOptions
// Tags: pipe delimited import TxtLoadOptions Aspose.Cells | custom separator loading txt C# | convert txt to xlsx Aspose.Cells | Aspose.Cells workbook save as xlsx | C# load text with custom delimiter Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsPipeDelimiterExample
{
    // The example checks for a pipe‑delimited text file, sets TxtLoadOptions.Separator to '|', loads the file into an Aspose.Cells Workbook, and then saves the workbook as an XLSX file, handling errors and reporting file paths.
    class Program
    {
        static void Main()
        {
            // Path to the pipe‑delimited source file
            string sourcePath = "data_pipe.txt";

            // Verify that the source file exists before attempting to load it
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {Path.GetFullPath(sourcePath)}");
                return;
            }

            try
            {
                // Configure load options to use pipe character as the separator
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Separator = '|'
                };

                // Load the pipe‑delimited file into a workbook using the configured options
                Workbook workbook = new Workbook(sourcePath, loadOptions);

                // Save the loaded workbook as an XLSX file
                string outputPath = "output.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
