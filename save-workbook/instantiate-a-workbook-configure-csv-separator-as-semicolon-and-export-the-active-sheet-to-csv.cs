// Title: Export the active worksheet to a semicolon‑delimited CSV file with Aspose.Cells for .NET
// AI Prompts: Write C# code that saves only the currently active worksheet of a Workbook as a CSV file using a semicolon as the field separator with Aspose.Cells. | Demonstrate how to set TxtSaveOptions.Separator to ';' and call Workbook.Save to generate a semicolon‑separated CSV from the active sheet in Aspose.Cells .NET.
// Common Searches: Aspose.Cells .NET how to export only the active sheet to a CSV with a custom delimiter | C# save worksheet as semicolon delimited CSV using Aspose.Cells | set CSV field separator to semicolon in Aspose.Cells TxtSaveOptions | export active worksheet to CSV without saving other sheets Aspose.Cells | configure Aspose.Cells to generate CSV files with ';' separator
// Tags: active worksheet CSV export Aspose.Cells | semicolon CSV delimiter TxtSaveOptions | save active sheet as CSV Aspose.Cells | custom CSV separator Aspose.Cells .NET | C# Aspose.Cells CSV export with custom delimiter

using Aspose.Cells;
using System;

// Creates a workbook, optionally fills cells, configures TxtSaveOptions.Separator to ';', and saves only the active worksheet as a semicolon‑delimited CSV file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Instantiate a new workbook (creates an empty workbook with one default worksheet)
        Workbook workbook = new Workbook();

        // Access the active worksheet (the first sheet by default)
        Worksheet activeSheet = workbook.Worksheets[workbook.Worksheets.ActiveSheetIndex];

        // (Optional) Populate some data in the active sheet for demonstration
        activeSheet.Cells["A1"].PutValue("Name");
        activeSheet.Cells["B1"].PutValue("Age");
        activeSheet.Cells["A2"].PutValue("John");
        activeSheet.Cells["B2"].PutValue(30);

        // Configure CSV save options to use semicolon as the field separator
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.CSV);
        csvOptions.Separator = ';';

        // Export the active sheet to a CSV file using the configured options
        workbook.Save("ActiveSheetExport.csv", csvOptions);
    }
}
