// Title: Add Row Numbers to First Column and Export Excel to CSV with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx workbook, inserts a leftmost column, fills it with sequential numbers starting at 1 for each data row, and saves the sheet directly as a CSV file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | Excel to CSV | prepend row numbers | insert column index 0 | SaveFormat.Csv | row numbering | worksheet manipulation | data export
// Common Searches: Aspose.Cells add row number column and save as CSV | C# insert column at beginning of worksheet before CSV export | How to prepend line numbers to CSV using Aspose.Cells | Convert Excel to CSV with index column .NET | Add sequential numbers to first column in Excel with Aspose
// Developer Intent: Insert a sequential index column at the left edge of the first worksheet and generate a CSV file.
// Use Cases: Create CSV logs with line numbers for audit trails. | Prepare data files for systems that require an explicit row identifier. | Facilitate manual data review by providing visible row numbers. | Export Excel sheets to CSV while preserving original row order for downstream processing.
// AI Prompts: Generate C# code using Aspose.Cells to add a leftmost row‑number column and save as CSV. | Show how to start numbering from a custom offset and optionally exclude header rows. | Demonstrate exporting only a specific worksheet to CSV after inserting row numbers. | Provide a LINQ‑based approach to fill the index column in one statement.

using System;
using Aspose.Cells;

namespace AsposeCellsCsvWithRowNumbers
{
    // Loads an .xlsx workbook, inserts a leftmost column, fills it with sequential numbers starting at 1 for each data row, and saves the sheet directly as a CSV file using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel workbook
            string sourcePath = "input.xlsx";

            // Load the workbook (lifecycle rule: create & load)
            Workbook workbook = new Workbook(sourcePath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Insert a new column at the very left (index 0)
            cells.InsertColumn(0);

            // Determine the last row that contains data
            int lastDataRow = cells.MaxDataRow;

            // Populate the new first column with sequential row numbers (starting at 1)
            for (int row = 0; row <= lastDataRow; row++)
            {
                cells[row, 0].PutValue(row + 1);
            }

            // Save the modified workbook as CSV (lifecycle rule: save)
            string csvPath = "output.csv";
            workbook.Save(csvPath, SaveFormat.Csv);
        }
    }
}
