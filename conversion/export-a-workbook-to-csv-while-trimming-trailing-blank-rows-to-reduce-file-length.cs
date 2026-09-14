// Title: Export an Excel workbook to CSV with Aspose.Cells for .NET while removing blank rows and trimming trailing cells
// AI Prompts: Generate C# code that loads an .xlsx file, removes all fully empty rows from the first worksheet, and saves the result as a CSV using Aspose.Cells with UTF-8 encoding and a comma separator. | Show how to configure TxtSaveOptions in Aspose.Cells to enable TrimTailingBlankCells, set the CSV separator, and specify the output encoding.
// Common Searches: Aspose.Cells how to delete empty rows before saving as CSV | C# export Excel to CSV without trailing blank cells using Aspose | Configure TxtSaveOptions for UTF-8 CSV output in Aspose.Cells .NET | Remove blank rows from worksheet then convert to CSV with Aspose.Cells
// Tags: Aspose.Cells delete empty rows | CSV export TrimTailingBlankCells option | TxtSaveOptions set CSV separator .NET | UTF-8 CSV generation Aspose.Cells | remove blank rows prior to CSV conversion

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Loads an Excel workbook, deletes completely blank rows from the first worksheet, configures TxtSaveOptions to trim trailing blank cells, use a comma separator and UTF-8 encoding, then saves the workbook as a compact CSV file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Remove all completely blank rows from the first worksheet
            // This reduces the length of the exported CSV by eliminating empty rows.
            workbook.Worksheets[0].Cells.DeleteBlankRows();

            // Configure CSV (text) save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Trim trailing blank cells in each row (optional, further reduces file size)
                TrimTailingBlankCells = true,
                // Use comma as the separator for CSV
                Separator = ',',
                // Set desired encoding
                Encoding = Encoding.UTF8
            };

            // Export the workbook to CSV using the configured options
            workbook.Save("output.csv", csvOptions);
        }
    }
}
