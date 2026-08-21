// Title: Import a 2D DateTime array and sort rows by earliest timestamp with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, imports a two‑dimensional DateTime matrix starting at A1, adds a helper column that holds each row's minimum timestamp, configures DataSorter to sort the range (no header row) in ascending order, optionally removes the helper column, and saves the result as SortedTimestamps.xlsx.
// Keywords: Aspose.Cells ImportTwoDimensionArray | C# sort rows by minimum timestamp | DataSorter helper column | Excel sort by calculated column .NET | DateTime matrix import Aspose.Cells | Aspose.Cells row sorting example
// Common Searches: how to import a 2D DateTime array into Aspose.Cells | sort worksheet rows by earliest date per row Aspose.Cells | using a helper column with DataSorter in C# | remove temporary column after sorting Aspose.Cells
// Developer Intent: Load a DateTime matrix into a worksheet and order the rows based on the earliest timestamp in each row.
// Use Cases: Arrange schedule data where each row contains multiple start times and the earliest start should appear first. | Generate an events log with several date columns and need rows ordered by the earliest event date. | Prepare data for timeline visualizations by sorting rows on the minimum timestamp across columns.
// AI Prompts: Generate C# code that uses Aspose.Cells to import a 2D DateTime array, adds a helper column with each row's minimum timestamp, sorts rows ascending, deletes the helper column, and saves the workbook. | Explain step‑by‑step how to set up Aspose.Cells DataSorter to sort rows based on a calculated column when the sheet has no header row. | Show an alternative LINQ‑based approach to compute row order by earliest timestamp and write the sorted rows back to the worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsTimestampSorting
{
    // Creates a workbook, imports a two‑dimensional DateTime matrix starting at A1, adds a helper column that holds each row's minimum timestamp, configures DataSorter to sort the range (no header row) in ascending order, optionally removes the helper column, and saves the result as SortedTimestamps.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Define a two‑dimensional array of timestamps (DateTime objects)
            object[,] timestamps = new object[,]
            {
                { new DateTime(2023, 5, 10, 14, 30, 0), new DateTime(2023, 5, 12, 9, 15, 0), new DateTime(2023, 5, 11, 18, 45, 0) },
                { new DateTime(2023, 4, 20, 8, 0, 0),  new DateTime(2023, 4, 22, 16, 0, 0), new DateTime(2023, 4, 21, 12, 30, 0) },
                { new DateTime(2023, 6, 1, 10, 0, 0),   new DateTime(2023, 6, 3, 11, 0, 0),  new DateTime(2023, 6, 2, 9, 0, 0) }
            };

            int rowCount = timestamps.GetLength(0);
            int colCount = timestamps.GetLength(1);

            // Import the two‑dimensional array into the worksheet starting at cell A1 (row 0, column 0)
            cells.ImportTwoDimensionArray(timestamps, 0, 0);

            // Add a helper column that stores the earliest timestamp of each row
            int helperColumnIndex = colCount; // place it right after the imported data
            for (int i = 0; i < rowCount; i++)
            {
                DateTime earliest = DateTime.MaxValue;
                for (int j = 0; j < colCount; j++)
                {
                    DateTime current = (DateTime)timestamps[i, j];
                    if (current < earliest)
                        earliest = current;
                }
                // Put the earliest timestamp into the helper column
                cells[i, helperColumnIndex].PutValue(earliest);
            }

            // Configure the DataSorter to sort rows based on the helper column (ascending)
            DataSorter sorter = workbook.DataSorter;
            sorter.HasHeaders = false;               // No header row in this data set
            sorter.AddKey(helperColumnIndex, SortOrder.Ascending);

            // Perform the sort on the full range including the helper column
            sorter.Sort(cells, 0, 0, rowCount - 1, helperColumnIndex);

            // (Optional) Remove the helper column after sorting
            // cells.DeleteColumn(helperColumnIndex);

            // Save the workbook with the sorted data
            workbook.Save("SortedTimestamps.xlsx");
        }
    }
}
