using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class SafeRowCopyDemo
    {
        // Copies rows from source to destination while ensuring the destination does not exceed the worksheet row limit.
        public static void SafeCopyRows(Cells sourceCells, int sourceRowIndex,
                                        Cells destinationCells, int destinationRowIndex,
                                        int rowNumber, int maxRowIndex)
        {
            try
            {
                // Calculate the last row index that would be written after the copy.
                int lastDestinationRow = destinationRowIndex + rowNumber - 1;

                // If the copy would exceed the worksheet limit, adjust the number of rows to copy.
                if (lastDestinationRow > maxRowIndex)
                {
                    // Determine how many rows can actually be copied.
                    int allowedRows = maxRowIndex - destinationRowIndex + 1;

                    // If no rows can be copied, simply return without invoking CopyRows to avoid an exception.
                    if (allowedRows <= 0)
                    {
                        Console.WriteLine("Destination start row is beyond the worksheet row limit. No rows were copied.");
                        return;
                    }

                    Console.WriteLine($"Requested copy exceeds row limit. Adjusting row count from {rowNumber} to {allowedRows}.");
                    rowNumber = allowedRows;
                }

                // Perform the copy operation using Aspose.Cells API.
                destinationCells.CopyRows(sourceCells, sourceRowIndex, destinationRowIndex, rowNumber);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during row copy: {ex.Message}");
            }
        }

        public static void Run()
        {
            try
            {
                // Create source workbook and populate some rows.
                Workbook sourceWb = new Workbook();
                Worksheet sourceSheet = sourceWb.Worksheets[0];
                Cells srcCells = sourceSheet.Cells;
                for (int i = 0; i < 10; i++)
                {
                    srcCells[i, 0].PutValue($"Source Row {i + 1}");
                }

                // Create destination workbook.
                Workbook destWb = new Workbook();
                Worksheet destSheet = destWb.Worksheets[0];
                Cells destCells = destSheet.Cells;

                // Example: attempt to copy 5 rows starting at row index near the limit for XLS format.
                int sourceStart = 0;
                int destinationStart = destWb.Settings.MaxRow - 2; // e.g., 65533 for XLS
                int rowsToCopy = 5;

                // Ensure the destination workbook file (if loading from disk) exists – not needed here but kept for completeness.
                // if (!File.Exists("template.xlsx")) { /* handle missing file */ }

                // Perform safe copy.
                SafeCopyRows(srcCells, sourceStart, destCells, destinationStart, rowsToCopy, destWb.Settings.MaxRow);

                // Save the result.
                string outputPath = "SafeCopyRowsOutput.xlsx";
                destWb.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application.
    internal class Program
    {
        private static void Main(string[] args)
        {
            SafeRowCopyDemo.Run();
        }
    }
}