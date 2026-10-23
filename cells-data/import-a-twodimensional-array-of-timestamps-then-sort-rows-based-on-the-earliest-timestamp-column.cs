// Title: Import a two‑dimensional DateTime array into an Aspose.Cells worksheet and sort rows by the earliest timestamp in C#
// Common Searches: C# Aspose.Cells import DateTime[,] array into worksheet | Sort worksheet rows by earliest date column with Aspose.Cells | Aspose.Cells ImportTwoDimensionArray with DateTime values and order rows | Clear and rewrite worksheet data after sorting timestamps in Aspose.Cells C# | Determine minimum timestamp per row and sort rows in Aspose.Cells workbook
// Tags: ImportTwoDimensionArray DateTime Aspose.Cells | order rows by minimum DateTime Aspose.Cells | clear worksheet range Aspose.Cells | write sorted data to worksheet Aspose.Cells | C# Aspose.Cells timestamp sorting

using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Cells;

namespace AsposeCellsTimestampSorting
{
    // The example creates a DateTime[,] array, converts it to object[,] for ImportTwoDimensionArray, imports the data into the first worksheet, computes the earliest timestamp in each row, sorts the rows ascending by that timestamp, clears the original range, writes the sorted rows back, and saves the workbook as SortedTimestamps.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Sample two‑dimensional array of timestamps (DateTime)
                DateTime[,] timestamps = new DateTime[,]
                {
                    { new DateTime(2023, 5, 10, 14, 30, 0), new DateTime(2023, 5, 12, 9, 15, 0), new DateTime(2023, 5, 11, 18, 45, 0) },
                    { new DateTime(2023, 4, 20, 8, 0, 0),  new DateTime(2023, 4, 22, 10, 30, 0), new DateTime(2023, 4, 21, 16, 0, 0) },
                    { new DateTime(2023, 6, 1, 12, 0, 0),  new DateTime(2023, 6, 3, 14, 0, 0),  new DateTime(2023, 6, 2, 9, 0, 0) }
                };

                // Convert DateTime[,] to object[,] because ImportTwoDimensionArray expects object[,]
                int rows = timestamps.GetLength(0);
                int cols = timestamps.GetLength(1);
                object[,] objArray = new object[rows, cols];
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        objArray[r, c] = timestamps[r, c];
                    }
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Import the two‑dimensional array starting at cell A1 (row 0, column 0)
                sheet.Cells.ImportTwoDimensionArray(objArray, 0, 0);

                // Determine the used range dimensions
                int totalRows = sheet.Cells.MaxDataRow + 1;      // rows count
                int totalCols = sheet.Cells.MaxDataColumn + 1;   // columns count

                // Build a list of rows with their earliest timestamp
                var rowsWithMin = new List<(int RowIndex, DateTime MinTimestamp, object[] Values)>();

                for (int r = 0; r < totalRows; r++)
                {
                    var rowValues = new object[totalCols];
                    DateTime min = DateTime.MaxValue;

                    for (int c = 0; c < totalCols; c++)
                    {
                        var cell = sheet.Cells[r, c];
                        object val = cell.Value;
                        rowValues[c] = val;

                        if (val is DateTime dt && dt < min)
                        {
                            min = dt;
                        }
                    }

                    rowsWithMin.Add((r, min, rowValues));
                }

                // Sort rows by the earliest timestamp (ascending)
                var sortedRows = rowsWithMin.OrderBy(item => item.MinTimestamp).ToList();

                // Clear existing data in the used range
                sheet.Cells.ClearContents(0, 0, totalRows, totalCols);

                // Write sorted rows back to the worksheet
                for (int r = 0; r < sortedRows.Count; r++)
                {
                    var values = sortedRows[r].Values;
                    for (int c = 0; c < totalCols; c++)
                    {
                        sheet.Cells[r, c].Value = values[c];
                    }
                }

                // Save the workbook to a file
                workbook.Save("SortedTimestamps.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
