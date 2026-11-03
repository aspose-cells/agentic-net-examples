// Title: Convert a pipe‑delimited TXT file to XLSX using Aspose.Cells TxtLoadOptions in C#
// Common Searches: C# Aspose.Cells load pipe delimited text with custom separator and export to Excel | How to set TxtLoadOptions.Separator to '|' for TXT to XLSX conversion in .NET | Example of converting pipe separated values file to .xlsx using Aspose.Cells library | Load a text file with pipe delimiter into Workbook and save as XLSX in C# | Aspose.Cells TxtLoadOptions pipe character usage for Excel export
// Tags: Aspose.Cells TxtLoadOptions pipe separator | C# load pipe delimited TXT into Workbook | convert pipe separated values to XLSX Aspose | custom delimiter import with Aspose.Cells | save Aspose.Cells workbook as XLSX | TxtLoadOptions Separator property usage

using System;
using Aspose.Cells;

namespace PipeDelimitedToXlsx
{
    // The code sets TxtLoadOptions.Separator to the pipe character, loads a pipe‑delimited TXT file into an Aspose.Cells Workbook, and saves the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the pipe‑delimited source file
            string sourcePath = "input.txt";

            // Path for the resulting XLSX file
            string destinationPath = "output.xlsx";

            // Configure load options to use pipe (|) as the column separator
            TxtLoadOptions loadOptions = new TxtLoadOptions
            {
                Separator = '|'
            };

            // Load the pipe‑delimited file into a workbook using the specified options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook in XLSX format
            workbook.Save(destinationPath, SaveFormat.Xlsx);
        }
    }
}
