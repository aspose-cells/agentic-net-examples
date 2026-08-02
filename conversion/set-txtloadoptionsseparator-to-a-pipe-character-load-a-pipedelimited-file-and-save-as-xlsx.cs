// Title: C# – Convert Pipe‑Delimited Text to XLSX with Aspose.Cells TxtLoadOptions.Separator
// Description: Shows how to configure TxtLoadOptions.Separator to ‘|’, import a pipe‑separated .txt file into an Aspose.Cells Workbook, and save the result as an .xlsx workbook, with sample data creation and basic exception handling.
// Keywords: Aspose.Cells | C# | pipe delimited | TxtLoadOptions | Separator | convert txt to xlsx | custom delimiter | load text file | save as xlsx | Excel conversion
// Common Searches: Aspose.Cells set separator to pipe | load pipe delimited file C# | convert txt to xlsx using Aspose | TxtLoadOptions custom delimiter example | C# code for pipe separated values to Excel
// Developer Intent: Import a pipe‑separated text file into a Workbook via TxtLoadOptions.Separator and export it as an XLSX document.
// Use Cases: Turn log files that use ‘|’ as column delimiter into searchable Excel reports. | Migrate legacy pipe‑separated export files into modern .xlsx dashboards. | Provide business users with an Excel view of configuration data stored in plain‑text files.
// AI Prompts: Write C# code that reads a pipe‑separated .txt file with Aspose.Cells TxtLoadOptions and saves it as .xlsx, including error handling. | Explain how to change TxtLoadOptions.Separator for other delimiters such as tab, comma, or semicolon. | Create a step‑by‑step guide to batch‑process multiple pipe‑delimited files into separate XLSX workbooks using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace PipeDelimitedConversion
{
    // Shows how to configure TxtLoadOptions.Separator to ‘|’, import a pipe‑separated .txt file into an Aspose.Cells Workbook, and save the result as an .xlsx workbook, with sample data creation and basic exception handling.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input_pipe.txt";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists; create a simple sample if missing
                if (!File.Exists(inputPath))
                {
                    File.WriteAllText(inputPath, "Column1|Column2|Column3\nValue1|Value2|Value3");
                }

                // Configure load options to use pipe as the column separator
                TxtLoadOptions loadOptions = new TxtLoadOptions
                {
                    Separator = '|'
                };

                // Load the pipe‑delimited text file into a workbook
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Save the workbook in XLSX format
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Conversion successful. Output saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
