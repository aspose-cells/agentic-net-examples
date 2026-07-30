// Title: C# – Save Workbook as CSV with Blank‑Row Separators using Aspose.Cells TxtSaveOptions
// Description: Shows how to build a workbook, insert values with empty rows, set TxtSaveOptions (UTF‑8, comma delimiter, KeepSeparatorsForBlankRow = true) and save it as a CSV file that retains column separators on completely blank rows.
// Keywords: Aspose.Cells | C# CSV export | TxtSaveOptions | KeepSeparatorsForBlankRow | blank rows CSV | UTF-8 CSV | comma delimiter | preserve separators | export workbook to CSV | conversion
// Common Searches: Aspose.Cells keep commas on empty rows | TxtSaveOptions KeepSeparatorsForBlankRow C# | Save CSV with blank rows Aspose | CSV export UTF-8 Aspose.Cells .NET | preserve column count in CSV using Aspose
// Developer Intent: Export a workbook to CSV while ensuring that rows without data still contain the appropriate delimiters.
// Use Cases: Generate CSV reports that include visual spacing rows without breaking column alignment. | Feed CSV files into legacy parsers that require a fixed number of columns per line. | Create multilingual CSV outputs where empty rows must keep delimiter structure. | Automate data pipelines where blank rows act as section markers but column consistency is mandatory.
// AI Prompts: Provide a C# example that uses Aspose.Cells TxtSaveOptions to save a workbook as CSV with KeepSeparatorsForBlankRow enabled and UTF‑8 encoding. | Describe the impact of the KeepSeparatorsForBlankRow property on CSV output and suggest a method to verify the separators on empty rows. | Modify the code to use a tab character as the CSV delimiter while still preserving separators on blank rows.

using System;
using System.Text;
using Aspose.Cells;

// Shows how to build a workbook, insert values with empty rows, set TxtSaveOptions (UTF‑8, comma delimiter, KeepSeparatorsForBlankRow = true) and save it as a CSV file that retains column separators on completely blank rows.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Add data with blank rows in between
        cells[0, 0].PutValue("a");
        cells[0, 1].PutValue("b");
        // rows 1 and 2 remain blank
        cells[3, 0].PutValue("c");
        cells[4, 1].PutValue("d");

        // Configure CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions
        {
            Encoding = Encoding.UTF8,   // use UTF‑8 encoding
            Separator = ',',           // comma delimiter for CSV
            KeepSeparatorsForBlankRow = true // keep separators on blank rows
        };

        // Save the workbook as CSV using the custom options
        workbook.Save("output.csv", csvOptions);
    }
}
