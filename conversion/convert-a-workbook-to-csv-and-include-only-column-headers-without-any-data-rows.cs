// Title: Export only the first row (column headers) of an Excel worksheet to CSV using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that loads an Excel workbook, clears all data rows while preserving the first row, and writes the result to a CSV file. | Demonstrate how to configure CSV save options in Aspose.Cells to output only the worksheet's column header row.
// Common Searches: Aspose.Cells C# export worksheet header row to CSV without data rows | How to save only the first row of an Excel sheet as CSV using Aspose.Cells | Remove data rows before converting Excel to CSV with Aspose.Cells .NET
// Tags: Aspose.Cells header-only CSV conversion | delete rows after header Aspose.Cells C# | configure CSV save options Aspose.Cells | export worksheet column names as CSV Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

// The example loads an Excel workbook, deletes every row below the first (header) row in the first worksheet, and saves the remaining header row as a CSV file using Aspose.Cells with appropriate CSV save options.
class WorkbookToCsvHeadersOnly
{
    static void Main()
    {
        // Path to the source workbook (any supported format, e.g., XLSX)
        string sourcePath = "input.xlsx";

        // Path for the resulting CSV file containing only the header row
        string csvPath = "output.csv";

        // Load the workbook with default load options
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook workbook = new Workbook(sourcePath, loadOptions);

        // Work with the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Determine the last row that contains data
        int lastDataRow = sheet.Cells.MaxDataRow;

        // If there are rows beyond the header (row 0), delete them
        if (lastDataRow > 0)
        {
            // Delete rows starting from row index 1 (second row) to the last data row
            sheet.Cells.DeleteRows(1, lastDataRow);
        }

        // Prepare CSV save options
        TxtSaveOptions csvOptions = new TxtSaveOptions(SaveFormat.Csv);
        // Ensure leading blank rows/columns are trimmed (optional but typical for CSV)
        csvOptions.TrimLeadingBlankRowAndColumn = true;

        // Save the workbook as CSV; only the header row remains
        workbook.Save(csvPath, csvOptions);
    }
}
