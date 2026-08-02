using System;
using Aspose.Cells;

namespace AsposeCellsTimestampSorting
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Two‑dimensional array of timestamps (DateTime objects)
            object[,] timestamps = new object[,]
            {
                { new DateTime(2023, 10, 5, 10, 0, 0), new DateTime(2023, 10, 5, 12, 0, 0) },
                { new DateTime(2023, 9, 30, 9, 30, 0), new DateTime(2023, 10, 1, 8, 0, 0) },
                { new DateTime(2023, 10, 2, 14, 0, 0), new DateTime(2023, 10, 2, 16, 0, 0) }
            };

            // Import the 2‑D array into the worksheet starting at cell A1 (row 0, column 0)
            cells.ImportTwoDimensionArray(timestamps, 0, 0);

            // Determine the range that contains the imported data
            int startRow = 0;
            int startColumn = 0;
            int endRow = timestamps.GetLength(0) - 1;      // last row index
            int endColumn = timestamps.GetLength(1) - 1;   // last column index

            // Configure the DataSorter to sort rows by the first column (earliest timestamp)
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = false;                     // No header row in this data
            sorter.AddKey(0, SortOrder.Ascending);         // Sort by column A (index 0)

            // Perform the sort on the defined area (top‑to‑bottom)
            sorter.Sort(cells, startRow, startColumn, endRow, endColumn);

            // Save the workbook to a file
            workbook.Save("SortedTimestamps.xlsx");
        }
    }
}