// Title: Import a 2D DateTime array into an Aspose.Cells worksheet, sort rows by the earliest timestamp, and save the workbook (C#)
// Common Searches: c# aspose.cells import DateTime[,] array into worksheet | aspose.cells sort Excel rows by minimum date in each row c# | move sorted rows to a new sheet with Aspose.Cells and save as xlsx | remove original worksheet after sorting data with Aspose.Cells C# | save workbook after sorting timestamps using Aspose.Cells .NET
// Tags: load DateTime matrix into worksheet Aspose.Cells | row sorting based on earliest timestamp Aspose.Cells | add new worksheet for sorted rows Aspose.Cells | delete source worksheet after row sorting Aspose.Cells | export workbook to XLSX using Aspose.Cells C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace TimestampRowSorter
{
    // Demonstrates how to load a two‑dimensional DateTime array into an Aspose.Cells worksheet, compute the earliest timestamp per row, reorder rows ascending by that timestamp, copy the sorted rows to a newly created worksheet, optionally delete the original sheet, and finally save the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // 1. Define a two‑dimensional array of timestamps (DateTime values)
                DateTime[,] timestamps = new DateTime[,]
                {
                    { new DateTime(2023, 5, 10, 14, 30, 0), new DateTime(2023, 5, 12, 9, 15, 0), new DateTime(2023, 5, 11, 18, 45, 0) },
                    { new DateTime(2023, 4, 20, 8, 0, 0),  new DateTime(2023, 4, 22, 16, 30, 0), new DateTime(2023, 4, 21, 12, 0, 0) },
                    { new DateTime(2023, 6, 1, 10, 0, 0),  new DateTime(2023, 6, 3, 11, 30, 0), new DateTime(2023, 6, 2, 9, 45, 0) },
                    { new DateTime(2023, 3, 15, 7, 20, 0), new DateTime(2023, 3, 16, 14, 55, 0), new DateTime(2023, 3, 14, 20, 10, 0) }
                };

                // 2. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                int rowCount = timestamps.GetLength(0);
                int colCount = timestamps.GetLength(1);

                // 3. Import the array into the worksheet cells
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < colCount; j++)
                    {
                        // PutValue automatically detects DateTime and stores it as a date/time cell
                        sheet.Cells[i, j].PutValue(timestamps[i, j]);
                    }
                }

                // 4. Determine the earliest timestamp for each row
                var rowInfo = new List<(DateTime min, int index)>();
                for (int i = 0; i < rowCount; i++)
                {
                    DateTime min = timestamps[i, 0];
                    for (int j = 1; j < colCount; j++)
                    {
                        if (timestamps[i, j] < min)
                            min = timestamps[i, j];
                    }
                    rowInfo.Add((min, i));
                }

                // 5. Sort rows by the earliest timestamp (ascending)
                rowInfo.Sort((a, b) => a.min.CompareTo(b.min));

                // 6. Create a new worksheet to hold sorted rows
                Worksheet sortedSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                sortedSheet.Name = "SortedData";

                // 7. Copy rows to the new worksheet in sorted order
                for (int destRow = 0; destRow < rowCount; destRow++)
                {
                    int srcRow = rowInfo[destRow].index;
                    for (int col = 0; col < colCount; col++)
                    {
                        sortedSheet.Cells[destRow, col].PutValue(timestamps[srcRow, col]);
                    }
                }

                // 8. Optionally remove the original unsorted sheet
                workbook.Worksheets.RemoveAt(0);

                // 9. Save the workbook to a file (you can change the path as needed)
                string outputPath = "SortedTimestamps.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Rows sorted successfully. Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the timestamps:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
