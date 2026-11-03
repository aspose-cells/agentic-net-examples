// Title: Set a semicolon delimiter with TxtLoadOptions and convert a CSV file to XLSX using Aspose.Cells for .NET
// Common Searches: how to load a semicolon‑delimited CSV in C# with Aspose.Cells TxtLoadOptions | convert CSV file with custom separator to XLSX using Aspose.Cells .NET | Aspose.Cells example for setting TxtLoadOptions.Separator to ';' and saving as XLSX | C# code to read CSV using semicolon as delimiter and export to Excel workbook
// Tags: TxtLoadOptions custom separator CSV | CSV to XLSX conversion Aspose.Cells | semicolon delimited CSV import C# | save workbook as Xlsx Aspose.Cells | load CSV with TxtLoadOptions Aspose.Cells

using System;
using Aspose.Cells;

// The sample configures TxtLoadOptions.Separator to a semicolon, loads a semicolon‑delimited CSV into a Workbook, and then saves the workbook as an XLSX file using Aspose.Cells for .NET.
class CsvToXlsxConverter
{
    static void Main()
    {
        // Specify the path to the source CSV file
        string csvPath = "input.csv";

        // Configure load options for CSV: set the separator to semicolon
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.CSV);
        loadOptions.Separator = ';';

        // Load the CSV file into a workbook using the configured options
        Workbook workbook = new Workbook(csvPath, loadOptions);

        // Specify the path for the output XLSX file
        string xlsxPath = "output.xlsx";

        // Save the workbook in XLSX format
        workbook.Save(xlsxPath, SaveFormat.Xlsx);
    }
}
