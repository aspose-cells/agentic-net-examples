// Title: Apply a custom numeric format to a cell and export the workbook as a tab‑delimited CSV (TSV) using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# apply custom number format before exporting to TSV | Save an Excel workbook as a tab‑separated CSV using Aspose.Cells .NET | How to set a custom numeric style and generate a tab delimited CSV with Aspose.Cells | C# example of TxtSaveOptions with tab separator for CSV export | Convert .xlsx to .tsv with custom formatting using Aspose.Cells library
// Tags: Aspose.Cells custom numeric style | TxtSaveOptions tab separator CSV | TSV export using Aspose.Cells | cell formatting API C# | export workbook to tab delimited format .NET

using Aspose.Cells;
using System;

// The sample loads an existing XLSX workbook, applies a custom numeric format "#,##0.00" to cell A1, configures TxtSaveOptions with a tab separator, and saves the result as a tab‑delimited TSV file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: apply a custom number format to cell A1
        Cell cell = worksheet.Cells["A1"];
        cell.PutValue(1234.567); // sample numeric value

        // Define the custom number format (e.g., thousand separator with two decimals)
        Style style = cell.GetStyle();
        style.Custom = "#,##0.00";
        cell.SetStyle(style);

        // Configure CSV save options with a tab delimiter
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
        csvOptions.Separator = '\t';

        // Save the workbook as a tab‑delimited CSV file
        workbook.Save("output.tsv", csvOptions);
    }
}
