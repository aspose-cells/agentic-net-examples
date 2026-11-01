// Title: Set a semicolon delimiter with TxtLoadOptions and convert a CSV file to XLSX using Aspense.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set CSV separator to semicolon before loading workbook | Convert semicolon delimited CSV to XLSX using TxtLoadOptions in .NET | Load CSV with custom delimiter and save as Excel file with Aspose.Cells | C# example for TxtLoadOptions separator property with CSV files | How to change CSV delimiter to ';' when importing into Aspose.Cells workbook
// Tags: TxtLoadOptions semicolon delimiter CSV | Aspose.Cells CSV to XLSX conversion C# | custom CSV separator Aspose.Cells | load CSV with TxtLoadOptions .NET | save workbook as Xlsx Aspose.Cells

using System;
using Aspose.Cells;

// The code creates a TxtLoadOptions object for CSV, sets its Separator to a semicolon, loads the semicolon‑delimited CSV into a Workbook, and then saves the workbook as an XLSX file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Configure the CSV separator to a semicolon
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.CSV);
        loadOptions.Separator = ';';

        // Load the CSV file with the specified options
        Workbook workbook = new Workbook("input.csv", loadOptions);

        // Export the workbook to XLSX format
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
