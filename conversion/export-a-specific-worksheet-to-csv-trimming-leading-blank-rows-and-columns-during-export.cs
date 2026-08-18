// Title: C# – Export a Worksheet to CSV with Trimmed Leading Blank Rows/Columns (Aspose.Cells)
// Description: Loads an Excel workbook, selects a specific worksheet (e.g., the second sheet), and saves it as a CSV file using Aspose.Cells. The TxtSaveOptions are set to TrimLeadingBlankRowAndColumn = true and ExportAllSheets = false, ensuring only the active sheet is exported and any leading empty rows or columns are removed.
// Keywords: Aspose.Cells CSV export C# | trim leading blank rows Aspose.Cells | export single worksheet to CSV .NET | TxtSaveOptions TrimLeadingBlankRowAndColumn | save active sheet as CSV | Aspose.Cells conversion example
// Common Searches: Aspose.Cells export specific sheet to CSV | remove leading empty rows when saving CSV with Aspose.Cells | C# TxtSaveOptions CSV trim blank rows and columns | how to export only active worksheet to CSV using Aspose.Cells | Aspose.Cells CSV export without blank rows
// Developer Intent: Generate a CSV file from a chosen worksheet while automatically discarding leading blank rows and columns.
// Use Cases: Create clean CSV reports from a particular sheet in a multi‑sheet workbook. | Prepare data for systems that cannot process leading empty rows or columns. | Automate Excel‑to‑CSV conversion where only the active sheet is required.
// AI Prompts: Write C# code with Aspose.Cells to export the third worksheet of an Excel file to CSV, trimming leading blank rows and columns. | Show how to configure TxtSaveOptions so that only the active sheet is saved as CSV and all other sheets are ignored. | Explain the effect of TrimLeadingBlankRowAndColumn when saving a worksheet to CSV with Aspose.Cells.

using System;
using Aspose.Cells;

namespace ExportWorksheetToCsv
{
    // Loads an Excel workbook, selects a specific worksheet (e.g., the second sheet), and saves it as a CSV file using Aspose.Cells. The TxtSaveOptions are set to TrimLeadingBlankRowAndColumn = true and ExportAllSheets = false, ensuring only the active sheet is exported and any leading empty rows or columns are removed.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Optionally, set the worksheet you want to export as the active sheet
            // For example, export the second worksheet (index 1)
            workbook.Worksheets.ActiveSheetIndex = 1;

            // Configure CSV (text) save options
            TxtSaveOptions saveOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Trim leading blank rows and columns just like Excel does
                TrimLeadingBlankRowAndColumn = true,

                // Export only the active worksheet (default behavior, but set explicitly for clarity)
                ExportAllSheets = false
            };

            // Save the active worksheet to CSV with the specified options
            workbook.Save("output_trimmed.csv", saveOptions);
        }
    }
}
