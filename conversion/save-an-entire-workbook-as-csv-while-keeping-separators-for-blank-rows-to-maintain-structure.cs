// Title: Save an entire workbook as CSV with blank‑row separators using Aspose.Cells for .NET (C#)
// Description: Demonstrates how to export a full Aspose.Cells workbook to a CSV file while keeping column separators for empty rows. The example configures TxtSaveOptions with UTF‑8 encoding, a comma delimiter, and the KeepSeparatorsForBlankRow flag to preserve the original sheet layout in the CSV output.
// Keywords: Aspose.Cells CSV export | KeepSeparatorsForBlankRow | TxtSaveOptions CSV | C# export workbook to CSV | .NET save workbook as CSV | preserve empty rows CSV | CSV delimiter Aspose.Cells | UTF-8 CSV Aspose | Excel to CSV conversion .NET
// Common Searches: Aspose.Cells keep commas for blank rows CSV | TxtSaveOptions KeepSeparatorsForBlankRow C# example | Export whole workbook to CSV with empty rows | How to preserve blank rows when saving Excel as CSV using Aspose | CSV export options Aspose.Cells .NET
// Developer Intent: Export the complete workbook to a CSV file while retaining column separators for rows that contain no data.
// Use Cases: Generating CSV reports that maintain visual grouping by preserving blank rows from the source spreadsheet. | Providing data files to legacy systems that require a fixed column count per row, including placeholders for empty rows. | Converting spreadsheets with intermittent empty rows into CSV for downstream processing without losing row alignment.
// AI Prompts: Show a C# code snippet that saves an Aspose.Cells workbook as CSV with KeepSeparatorsForBlankRow enabled. | Explain how the KeepSeparatorsForBlankRow option affects CSV output and how to set encoding and delimiter with TxtSaveOptions. | Modify the example to use a semicolon as the separator and UTF‑16 encoding while still preserving blank rows.

using System;
using System.Text;
using Aspose.Cells;

// Demonstrates how to export a full Aspose.Cells workbook to a CSV file while keeping column separators for empty rows. The example configures TxtSaveOptions with UTF‑8 encoding, a comma delimiter, and the KeepSeparatorsForBlankRow flag to preserve the original sheet layout in the CSV output.
class SaveWorkbookAsCsvWithBlankRows
{
    static void Main()
    {
        // Create a new workbook and get its first worksheet's cells
        Workbook workbook = new Workbook();
        Cells cells = workbook.Worksheets[0].Cells;

        // Populate some data with intentional blank rows
        cells[0, 0].PutValue("a");
        cells[0, 1].PutValue("b");
        // rows 1 and 2 remain blank
        cells[3, 0].PutValue("c");
        cells[4, 1].PutValue("d");

        // Configure text save options for CSV
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Encoding = Encoding.UTF8,          // Use UTF-8 encoding
            Separator = ',',                  // Comma as delimiter
            KeepSeparatorsForBlankRow = true // Preserve separators for blank rows
        };

        // Save the entire workbook as CSV with the specified options
        workbook.Save("output.csv", saveOptions);
    }
}
