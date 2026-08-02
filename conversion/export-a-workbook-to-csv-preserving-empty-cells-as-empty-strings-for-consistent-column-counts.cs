// Title: Export a Workbook to CSV with Aspose.Cells (C#) – Preserve Empty Cells and Blank Rows
// Description: Shows how to save an Excel workbook as a CSV file using Aspose.Cells for .NET while writing delimiters for empty cells and completely blank rows. The code configures TxtSaveOptions (KeepSeparatorsForBlankRow, TrimLeadingBlankRowAndColumn = false, TrimTrailingBlankCells = false) and applies ASCII encoding to keep a uniform column count across all rows.
// Keywords: Aspose.Cells CSV export | C# Aspose.Cells TxtSaveOptions | KeepSeparatorsForBlankRow | TrimLeadingBlankRowAndColumn | TrimTrailingBlankCells | preserve empty cells CSV | blank rows CSV Aspose | fixed column count CSV | Aspose.Cells export options | CSV encoding Aspose
// Common Searches: Aspose.Cells keep empty columns when saving to CSV | How to retain blank rows in CSV using Aspose.Cells .NET | TxtSaveOptions KeepSeparatorsForBlankRow example | Prevent column trimming in Aspose.Cells CSV export | Export Excel to CSV with fixed columns Aspose | C# save workbook as CSV without removing empty cells
// Developer Intent: Create a CSV file from an Excel workbook that retains all empty fields and separator lines, ensuring each record maintains the same number of columns.
// Use Cases: Data exchange with legacy systems that require a fixed number of columns per row | Generating CSV templates where missing values must appear as empty strings | Producing reports that include visual separator rows without losing column alignment | Automating inventory or pricing feeds where some items lack certain attributes | Preparing CSV files for batch import tools that treat absent fields as empty cells
// AI Prompts: Write C# code with Aspose.Cells to export a workbook to CSV, keeping empty cells and blank rows unchanged. | Explain the impact of KeepSeparatorsForBlankRow, TrimLeadingBlankRowAndColumn, and TrimTrailingBlankCells on CSV output in Aspose.Cells. | Show how to set encoding and other TxtSaveOptions for a CSV export that preserves column count. | Provide a step‑by‑step guide to configure Aspose.Cells CSV export for fixed‑width column layouts. | Generate a sample CSV output illustrating preserved empty cells from the given workbook.

using System;
using System.Text;
using Aspose.Cells;

namespace AsposeCellsCsvExport
{
    // Shows how to save an Excel workbook as a CSV file using Aspose.Cells for .NET while writing delimiters for empty cells and completely blank rows. The code configures TxtSaveOptions (KeepSeparatorsForBlankRow, TrimLeadingBlankRowAndColumn = false, TrimTrailingBlankCells = false) and applies ASCII encoding to keep a uniform column count across all rows.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate data with intentional empty cells
            // Row 0
            cells[0, 0].PutValue("Item");
            cells[0, 1].PutValue("Price");
            // Row 1 - leave column B empty
            cells[1, 0].PutValue("Apple");
            // Row 2 - both columns have values
            cells[2, 0].PutValue("Banana");
            cells[2, 1].PutValue(1.25);
            // Row 3 - completely empty row (will be kept as empty separators)
            // Row 4 - only second column has value
            cells[4, 1].PutValue(0.99);

            // Configure CSV (text) save options
            TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv)
            {
                // Ensure separators are written for completely blank rows
                KeepSeparatorsForBlankRow = true,
                // Do not trim leading blank rows/columns so column count stays consistent
                TrimLeadingBlankRowAndColumn = false,
                // Keep trailing blank cells (default is false, set explicitly for clarity)
                TrimTailingBlankCells = false,
                // Use ASCII encoding for demonstration; adjust as needed
                Encoding = Encoding.ASCII
            };

            // Save the workbook as CSV with the configured options
            string outputPath = "output_preserve_empty_cells.csv";
            workbook.Save(outputPath, csvOptions);

            Console.WriteLine($"Workbook exported to CSV at: {outputPath}");
        }
    }
}
