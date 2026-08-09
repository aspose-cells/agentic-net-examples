// Title: C# – Convert Pipe‑Delimited Text to XLSX with Aspose.Cells
// Description: Demonstrates how to set TxtLoadOptions.Separator to the pipe character, load a pipe‑delimited .txt file into an Aspose.Cells Workbook, and save the result as an XLSX workbook.
// Keywords: Aspose.Cells | C# | pipe delimited | TxtLoadOptions | Separator | text to Excel | XLSX conversion
// Common Searches: Aspose.Cells set TxtLoadOptions separator to pipe | convert pipe delimited file to Excel .NET | load .txt with custom delimiter Aspose.Cells | save workbook as XLSX using C# | batch convert pipe separated files to XLSX
// Developer Intent: Load a pipe‑separated text file into a Workbook and export it as an XLSX file.
// Use Cases: Migrate legacy pipe‑separated reports into modern Excel dashboards. | Automate daily conversion of log files that use ‘|’ as a delimiter. | Integrate custom‑delimited data imports into a .NET analytics application.
// AI Prompts: Generate C# code that uses Aspose.Cells to read a pipe‑delimited .txt file by configuring TxtLoadOptions.Separator and saves it as .xlsx. | Explain how to configure TxtLoadOptions for any custom delimiter when converting text files to Excel with Aspose.Cells for .NET. | Provide a script to batch‑process a directory of ‘|’ delimited files into separate XLSX workbooks using Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsPipeDemo
{
    // Demonstrates how to set TxtLoadOptions.Separator to the pipe character, load a pipe‑delimited .txt file into an Aspose.Cells Workbook, and save the result as an XLSX workbook.
    class Program
    {
        static void Main()
        {
            // Path to the pipe‑delimited source file
            string sourcePath = "input_pipe.txt";

            // Create load options and set the pipe character as the separator
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.Separator = '|';

            // Load the pipe‑delimited file using the configured options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the loaded workbook as an XLSX file
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
