// Title: Use TxtLoadOptions to set a semicolon delimiter and convert a CSV file to XLSX with Aspose.Cells for .NET
// AI Prompts: Load a CSV file with a semicolon delimiter using TxtLoadOptions and save it as an XLSX workbook in C# with Aspose.Cells. | Configure a custom CSV separator, import the data into a Workbook, and export to Excel format programmatically using the Aspose.Cells .NET API.
// Common Searches: Aspose.Cells how to specify semicolon as CSV delimiter in C# | Convert CSV file that uses ';' as separator to Excel workbook using .NET | C# example for loading CSV with custom separator and saving as .xlsx with Aspose.Cells | TxtLoadOptions separator property usage for semicolon delimited CSV conversion
// Tags: TxtLoadOptions custom separator | CSV to XLSX conversion Aspose.Cells | custom CSV delimiter .NET | load CSV with Aspose.Cells | save workbook as Xlsx Aspose

using System;
using Aspose.Cells;

// The sample configures TxtLoadOptions.Separator to a semicolon, loads a semicolon‑delimited CSV file into an Aspose.Cells Workbook, and then saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        // Configure TxtLoadOptions to use semicolon as the CSV separator
        TxtLoadOptions loadOptions = new TxtLoadOptions();
        loadOptions.Separator = ';';

        // Path to the source CSV file (replace with actual file location)
        string csvFilePath = "input.csv";

        // Load the CSV file with the specified load options
        Workbook workbook = new Workbook(csvFilePath, loadOptions);

        // Export the loaded workbook to XLSX format
        string xlsxFilePath = "output.xlsx";
        workbook.Save(xlsxFilePath, SaveFormat.Xlsx);
    }
}
