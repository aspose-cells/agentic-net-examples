// Title: Load a UTF-16 encoded CSV containing Asian characters with TxtLoadOptions in Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# load CSV file with UTF-16 encoding for Chinese characters | Set Unicode (UTF-16) encoding when importing CSV using TxtLoadOptions in .NET | How to read a CSV with Asian language text in Aspose.Cells using TxtLoadOptions | Specify Encoding.Unicode for CSV load in Aspose.Cells C# example | Convert UTF-16 CSV to Excel workbook with Aspose.Cells .NET API
// Tags: TxtLoadOptions UTF-16 encoding for CSV | Aspose.Cells CSV Unicode import C# | LoadFormat.CSV encoding configuration | Workbook load CSV with Asian characters | C# Aspose.Cells set Encoding.Unicode

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates creating a TxtLoadOptions object for CSV format, setting its Encoding to Encoding.Unicode (UTF-16) to correctly read Asian characters, loading the CSV into a Workbook, and optionally saving it as an Excel file.
class Program
{
    static void Main()
    {
        // Specify load options for a CSV file with UTF-16 (Unicode) encoding
        TxtLoadOptions loadOptions = new TxtLoadOptions(LoadFormat.CSV);
        loadOptions.Encoding = Encoding.Unicode; // UTF-16

        // Load the CSV file using the defined options
        Workbook workbook = new Workbook("input.csv", loadOptions);

        // Save the workbook to verify successful loading (optional)
        workbook.Save("output.xlsx");
    }
}
