using System;
using Aspose.Cells;

namespace AsposeCellsKeywordCounter
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Load the workbook (create/load rule)
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define the column index to inspect (0 = column A)
            int columnIndex = 0;

            // Define the keyword to count
            string keyword = "Apple";

            // Initialize counter
            int occurrenceCount = 0;

            // Determine the range of rows that contain data
            int firstDataRow = cells.MinDataRow;
            int lastDataRow = cells.MaxDataRow;

            // Iterate through each cell in the specified column
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                // Retrieve the raw string value of the cell
                string cellValue = cells[row, columnIndex].StringValue;

                // Check if the cell contains the keyword (case‑sensitive)
                if (!string.IsNullOrEmpty(cellValue) && cellValue.Contains(keyword))
                {
                    occurrenceCount++;
                }
            }

            // Log the total count
            Console.WriteLine($"Total occurrences of \"{keyword}\" in column {CellsHelper.ColumnIndexToName(columnIndex)}: {occurrenceCount}");

            // Save the workbook (save rule) – optional, can be omitted if no changes are made
            workbook.Save("output.xlsx");
        }
    }
}