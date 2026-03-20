using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsFilterDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing workbook (replace with your actual file path)
            Workbook workbook = new Workbook("InputData.xlsx");

            // Access the worksheet to be filtered (by index or name)
            Worksheet worksheet = workbook.Worksheets[0]; // first worksheet

            // Define the range that contains the header row and data to be filtered
            // Example: columns A to B, rows 1 to 10 (Excel rows are 1‑based, zero‑based in API)
            worksheet.AutoFilter.Range = "A1:B10";

            // Apply a filter on the first column (field index 0) for a specific criteria
            // Change "TargetValue" to the value you want to filter by
            worksheet.AutoFilter.Filter(0, "TargetValue");

            // Refresh the filter to hide rows that do not meet the criteria
            worksheet.AutoFilter.Refresh();

            // Collect the rows that remain visible after filtering
            List<string[]> matchingRows = new List<string[]>();
            int maxRow = worksheet.Cells.MaxDataRow; // last row with data
            int maxCol = worksheet.Cells.MaxDataColumn; // last column with data

            // Start from row 1 to skip the header (row 0)
            for (int row = 1; row <= maxRow; row++)
            {
                // If the row is hidden, it does not match the filter criteria
                if (worksheet.Cells.Rows[row].IsHidden)
                    continue;

                // Extract cell values for the current row
                string[] values = new string[maxCol + 1];
                for (int col = 0; col <= maxCol; col++)
                {
                    values[col] = worksheet.Cells[row, col].StringValue;
                }
                matchingRows.Add(values);
            }

            // Output the matching rows to the console
            Console.WriteLine("Rows matching the filter criteria:");
            foreach (var rowValues in matchingRows)
            {
                Console.WriteLine(string.Join("\t", rowValues));
            }

            // Optionally, save the workbook with the applied filter for visual verification
            workbook.Save("FilteredOutput.xlsx");
        }
    }
}