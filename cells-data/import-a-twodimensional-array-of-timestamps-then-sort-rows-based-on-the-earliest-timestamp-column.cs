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
                { new DateTime(2023, 10, 1, 8, 0, 0), new DateTime(2023, 10, 1, 12, 0, 0) },
                { new DateTime(2023, 10, 2, 9, 30, 0), new DateTime(2023, 10, 2, 11, 15, 0) },
                { new DateTime(2023, 9, 30, 14, 0, 0), new DateTime(2023, 10, 1, 9, 0, 0) }
            };

            // Import the two‑dimensional array into the worksheet starting at cell A1 (row 0, column 0)
            cells.ImportTwoDimensionArray(timestamps, 0, 0);

            // Determine dimensions of the imported data
            int rowCount = timestamps.GetLength(0);
            int colCount = timestamps.GetLength(1);
            int helperCol = colCount; // column index for the temporary min‑timestamp column

            // Populate the helper column with the earliest timestamp of each row
            for (int r = 0; r < rowCount; r++)
            {
                DateTime min = DateTime.MaxValue;
                for (int c = 0; c < colCount; c++)
                {
                    DateTime current = (DateTime)timestamps[r, c];
                    if (current < min)
                        min = current;
                }
                cells[r, helperCol].PutValue(min);
            }

            // Configure the DataSorter to sort by the helper column (ascending)
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = false;               // No header row in this example
            sorter.AddKey(helperCol, SortOrder.Ascending);

            // Perform the sort on the full data range including the helper column
            sorter.Sort(cells, 0, 0, rowCount - 1, helperCol);

            // (Optional) Remove the helper column after sorting
            cells.DeleteColumn(helperCol);

            // Save the workbook
            workbook.Save("SortedTimestamps.xlsx");
        }
    }
}