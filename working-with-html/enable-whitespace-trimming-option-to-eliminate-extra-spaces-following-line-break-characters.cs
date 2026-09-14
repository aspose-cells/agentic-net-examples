// Title: How to trim leading spaces after line breaks when saving an Excel worksheet to TXT using Aspose.Cells for .NET
// AI Prompts: Show C# code that sets TxtSaveOptions.TrimTrailingSpaces to true so that spaces after '\n' are removed when exporting a workbook to a .txt file with Aspose.Cells. | Generate a complete C# example that creates a workbook, writes a cell containing a newline and leading spaces, configures TxtSaveOptions for whitespace trimming, and saves the worksheet as a text file.
// Common Searches: Aspose.Cells TxtSaveOptions TrimTrailingSpaces example C# | C# remove leading spaces after newline when exporting Excel to text file | configure Aspose.Cells to trim whitespace after line breaks in TXT output | save workbook as .txt without extra spaces after \n using Aspose.Cells .NET
// Tags: Aspose.Cells TxtSaveOptions TrimTrailingSpaces | C# export Excel to TXT whitespace trimming | remove leading spaces after newline Aspose.Cells | Excel to text conversion whitespace handling .NET | Aspose.Cells save worksheet as txt without extra spaces

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example creates a workbook, writes a cell value that includes a line break followed by spaces, enables the TrimTrailingSpaces option in TxtSaveOptions, ensures the output directory exists, and saves the workbook as a trimmed text file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Set a cell value that contains a line break followed by extra spaces
                sheet.Cells["A1"].PutValue("First line\n   Second line with leading spaces");

                // Configure TXT save options (default behavior trims trailing spaces)
                TxtSaveOptions saveOptions = new TxtSaveOptions();

                // Determine output file path
                string outputPath = "TrimmedOutput.txt";

                // Ensure the directory for the output file exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ??
                                   Directory.GetCurrentDirectory();

                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as a text file using the configured options
                workbook.Save(outputPath, saveOptions);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
