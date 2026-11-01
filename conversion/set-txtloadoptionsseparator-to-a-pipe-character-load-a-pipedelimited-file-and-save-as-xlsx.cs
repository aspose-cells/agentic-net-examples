// Title: Load a pipe‑delimited text file with TxtLoadOptions.Separator and export to XLSX using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load pipe delimited txt file with custom separator | How to convert pipe‑separated values to Excel workbook in .NET | Set TxtLoadOptions.Separator to '|' for reading pipe delimited files with Aspose.Cells | Save a workbook loaded from a pipe‑delimited text as .xlsx using Aspose.Cells | C# example for loading custom delimited text into Aspose.Cells and exporting to XLSX
// Tags: TxtLoadOptions custom delimiter | pipe delimited text to XLSX conversion | Aspose.Cells load txt with separator | C# export workbook to XLSX | Aspose.Cells delimiter handling

using System;
using Aspose.Cells;

// Demonstrates using Aspose.Cells for .NET to set TxtLoadOptions.Separator to the pipe character, load a pipe‑delimited text file into a Workbook, and save the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Define load options for a pipe‑delimited text file
        TxtLoadOptions loadOptions = new TxtLoadOptions
        {
            // Set the column separator to the pipe character '|'
            Separator = '|'
        };

        // Load the pipe‑delimited file into a workbook using the specified options
        Workbook workbook = new Workbook("input.txt", loadOptions);

        // Save the workbook in XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
