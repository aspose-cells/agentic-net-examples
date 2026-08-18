// Title: Export Workbook to CSV with Aspose.Cells – Remove Blank Rows and Trim Trailing Cells (C#)
// Description: The sample creates a workbook, fills a few cells, calls DeleteBlankRows, configures TxtSaveOptions (UTF‑8, TrimTailingBlankCells=true, KeepSeparatorsForBlankRow=false, ExportAllSheets=false), and saves a compact CSV that contains only the active worksheet.
// Keywords: Aspose.Cells CSV export | DeleteBlankRows | TrimTailingBlankCells | C# Aspose.Cells | remove empty rows CSV | export active worksheet | TxtSaveOptions | compact CSV | UTF-8 CSV Aspose
// Common Searches: Aspose.Cells delete blank rows before CSV | Trim trailing blank cells in CSV using Aspose | Save only active sheet to CSV Aspose.Cells C# | Prevent empty rows in CSV export Aspose | Configure TxtSaveOptions for CSV Aspose.Cells
// Developer Intent: Generate a CSV from a workbook that excludes all blank rows and trims trailing empty cells, using Aspose.Cells in C#.
// Use Cases: Reduce CSV size by eliminating trailing empty rows in generated reports. | Produce a clean UTF‑8 CSV without extra delimiters for downstream data pipelines. | Export only the active worksheet from a multi‑sheet workbook. | Ensure each CSV row contains only populated cells for reliable import.
// AI Prompts: Write C# code with Aspose.Cells to delete all blank rows and export the active worksheet to a CSV with TrimTailingBlankCells enabled. | Show how to set TxtSaveOptions for CSV: UTF‑8 encoding, KeepSeparatorsForBlankRow false, ExportAllSheets false. | Explain the impact of DeleteBlankRows and TrimTailingBlankCells on CSV file size and readability.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // The sample creates a workbook, fills a few cells, calls DeleteBlankRows, configures TxtSaveOptions (UTF‑8, TrimTailingBlankCells=true, KeepSeparatorsForBlankRow=false, ExportAllSheets=false), and saves a compact CSV that contains only the active worksheet.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["A2"].PutValue("Data1");
            cells["B2"].PutValue(100);
            cells["A3"].PutValue("Data2");
            cells["B3"].PutValue(200);

            // Add blank rows at the bottom (rows 4-6 are empty)
            // No need to put any values – they remain blank

            // Remove all blank rows (including trailing ones) to shorten the CSV
            cells.DeleteBlankRows();

            // Configure CSV (text) save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                Encoding = Encoding.UTF8,
                // Trim trailing blank cells in each row (optional, improves compactness)
                TrimTailingBlankCells = true,
                // Do not output separators for completely blank rows (default false)
                KeepSeparatorsForBlankRow = false,
                // Export only the active sheet
                ExportAllSheets = false
            };

            // Save the workbook as CSV using the configured options
            workbook.Save("output_trimmed.csv", csvOptions);
        }
    }
}
