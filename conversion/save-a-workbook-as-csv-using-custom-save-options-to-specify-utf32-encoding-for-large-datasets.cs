// Title: Save a large Aspose.Cells workbook as CSV with UTF‑32 encoding (C#)
// Description: Shows how to create a workbook, fill 10,000 rows, set TxtSaveOptions with Encoding.UTF32 and a comma separator, and save the file as a UTF‑32 encoded CSV using Aspose.Cells for .NET.
// Keywords: Aspose.Cells CSV UTF-32 | TxtSaveOptions C# | save workbook as CSV .NET | UTF-32 encoding Aspose.Cells | large dataset CSV export | comma delimiter Aspose.Cells | C# Excel to CSV conversion | custom save options Aspose.Cells
// Common Searches: Aspose.Cells save CSV with UTF-32 | C# export large Excel to CSV UTF-32 | TxtSaveOptions encoding example | set CSV separator Aspose.Cells | export 10000 rows to CSV using Aspose.Cells | UTF-32 CSV file Aspose.Cells .NET
// Developer Intent: Export a workbook to a CSV file using UTF‑32 encoding and a comma delimiter.
// Use Cases: Providing CSV output for systems that require UTF‑32 encoded text. | Generating CSV reports from massive Excel data while preserving all Unicode characters. | Creating locale‑independent CSV files for downstream processing pipelines. | Exporting data for big‑data workflows that expect UTF‑32 encoding.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to CSV with UTF‑32 encoding and a custom delimiter. | Explain how to stream a workbook directly to a CSV file using TxtSaveOptions to minimize memory usage. | Show how to modify the example to use UTF‑8 with BOM while keeping the comma separator. | Provide guidance on handling special characters when exporting to UTF‑32 CSV with Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

// Shows how to create a workbook, fill 10,000 rows, set TxtSaveOptions with Encoding.UTF32 and a comma separator, and save the file as a UTF‑32 encoded CSV using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Add header row
        sheet.Cells["A1"].PutValue("ID");
        sheet.Cells["B1"].PutValue("Value");

        // Populate a large dataset (example: 10,000 rows)
        for (int i = 0; i < 10000; i++)
        {
            sheet.Cells[i + 1, 0].PutValue(i + 1);               // ID column
            sheet.Cells[i + 1, 1].PutValue($"Data_{i + 1}");    // Value column
        }

        // Create CSV (text) save options
        TxtSaveOptions saveOptions = new TxtSaveOptions();

        // Specify UTF‑32 encoding for the output file
        saveOptions.Encoding = Encoding.UTF32;

        // Ensure the separator is a comma (CSV)
        saveOptions.Separator = ',';

        // Save the workbook as CSV using the custom options
        workbook.Save("large_dataset.csv", saveOptions);
    }
}
