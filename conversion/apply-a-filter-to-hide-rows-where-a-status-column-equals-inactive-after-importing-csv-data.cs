using System;
using Aspose.Cells;

namespace AsposeCellsFilterExample
{
    class Program
    {
        static void Main()
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Load the CSV data into a workbook
            Workbook workbook = new Workbook(csvPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range (including header row)
            int lastRow = sheet.Cells.MaxDataRow;          // zero‑based index of the last row with data
            int lastColumn = sheet.Cells.MaxDataColumn;    // zero‑based index of the last column with data

            // Build the address string for the autofilter range (e.g., "A1:D10")
            string startCell = "A1";
            string endCell = CellsHelper.CellIndexToName(lastRow, lastColumn);
            sheet.AutoFilter.Range = $"{startCell}:{endCell}";

            // Assume the "Status" column is the third column (C) – zero‑based index = 2
            int statusColumnIndex = 2;

            // Apply a filter to hide rows where Status = "Inactive"
            sheet.AutoFilter.Filter(statusColumnIndex, "Inactive");

            // Refresh the filter so that rows not matching the criteria become hidden
            sheet.AutoFilter.Refresh();

            // Save the filtered workbook
            workbook.Save("output.xlsx");
        }
    }
}