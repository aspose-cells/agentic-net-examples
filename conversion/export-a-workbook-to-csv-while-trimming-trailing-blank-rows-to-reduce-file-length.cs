// Title: C# – Export Aspose.Cells Workbook to CSV and Remove Blank Rows (Trim Trailing Cells)
// Description: Shows how to create a workbook, delete fully blank rows, and save it as a CSV using Aspose.Cells for .NET with UTF‑8 encoding, comma separator, and TrimTrailingBlankCells to produce a compact file.
// Keywords: Aspose.Cells | C# | .NET | CSV export | DeleteBlankRows | TrimTrailingBlankCells | TxtSaveOptions | SaveFormat.Csv | UTF-8 encoding | comma separator | remove empty rows | Excel to CSV
// Common Searches: Aspose.Cells export to CSV C# | DeleteBlankRows Aspose.Cells example | TrimTrailingBlankCells CSV Aspose | how to remove blank rows before CSV export Aspose | save workbook as CSV without empty rows .NET
// Developer Intent: Create a CSV file from an Excel workbook while automatically discarding completely empty rows and trimming trailing empty cells to minimize file size.
// Use Cases: Generate clean CSV reports from spreadsheets that contain intermittent blank rows. | Prepare data for import into systems that reject rows with no values. | Reduce CSV size for large datasets by eliminating trailing blank cells. | Automate data‑export pipelines in .NET applications using Aspose.Cells.
// AI Prompts: Write C# code using Aspose.Cells to delete blank rows and export to CSV with TrimTrailingBlankCells enabled. | Explain the difference between DeleteBlankRows and TrimTrailingBlankCells when saving a workbook as CSV. | Show how to configure TxtSaveOptions for UTF‑8 CSV with a custom delimiter in Aspose.Cells.

using System;
using System.Text;
using Aspose.Cells;

// Shows how to create a workbook, delete fully blank rows, and save it as a CSV using Aspose.Cells for .NET with UTF‑8 encoding, comma separator, and TrimTrailingBlankCells to produce a compact file.
class ExportCsvTrimTrailingRows
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some data with intentional blank rows
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue("Row 1");
        // Row 3 is left blank
        cells["A4"].PutValue("Row 2");
        // Row 5 is left blank
        cells["A6"].PutValue("Row 3");

        // Remove all completely blank rows to eliminate trailing empty rows
        cells.DeleteBlankRows();

        // Configure CSV (text) save options
        TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
        {
            Encoding = Encoding.UTF8,          // Use UTF‑8 encoding
            Separator = ',',                  // Comma as CSV separator
            TrimTailingBlankCells = true      // Trim trailing blank cells in each row
        };

        // Save the workbook as a CSV file with the specified options
        workbook.Save("output.csv", saveOptions);
    }
}
