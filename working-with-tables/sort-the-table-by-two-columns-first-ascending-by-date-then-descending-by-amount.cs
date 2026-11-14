// Title: Sort an Excel table by date (ascending) and amount (descending) using Aspose.Cells in C#
// Common Searches: Aspose.Cells C# sort worksheet rows by date column ascending then amount column descending | How to sort Excel data with multiple keys using Aspose.Cells .NET | C# code to sort a range A1:C100 by date and then by numeric value with Aspose.Cells | Programmatically order Excel rows by earliest date and highest amount using Aspose.Cells | Custom comparator for multi‑column sorting in Aspose.Cells C# example
// Tags: Aspose.Cells multi‑column sorting C# | Excel date ascending amount descending sort | C# custom row comparator Aspose.Cells | sorting worksheet range programmatically .NET | Aspose.Cells sort by date and numeric column

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an input.xlsx workbook, extracts rows from the A1:C100 range (excluding the header), sorts them first by the Date column in ascending order and then by the Amount column in descending order using a custom List<object[]> comparator, writes the sorted data back to the worksheet, and saves the result as output.xlsx, with checks for missing files and empty data.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook.
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Define the data range (A1:C100). Adjust if the actual data size differs.
            const int startRow = 0;      // Header row index
            const int startColumn = 0;   // Column A
            const int totalColumns = 3;  // A, B, C

            // Determine the last row that contains data within the defined range.
            int lastRow = Math.Min(sheet.Cells.MaxDataRow, startRow + 99);
            if (lastRow <= startRow)
            {
                Console.WriteLine("No data rows found to sort.");
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
                return;
            }

            // Read data rows (excluding header) into a list.
            var rows = new List<object[]>();
            for (int r = startRow + 1; r <= lastRow; r++)
            {
                var rowValues = new object[totalColumns];
                for (int c = 0; c < totalColumns; c++)
                {
                    rowValues[c] = sheet.Cells[r, startColumn + c].Value;
                }
                rows.Add(rowValues);
            }

            // Perform sorting: first by Date (column 0) ascending, then by Amount (column 1) descending.
            rows.Sort((a, b) =>
            {
                // Compare dates (handle nulls).
                DateTime dateA = a[0] is DateTime da ? da : DateTime.MinValue;
                DateTime dateB = b[0] is DateTime db ? db : DateTime.MinValue;
                int dateComp = dateA.CompareTo(dateB);
                if (dateComp != 0) return dateComp;

                // Compare amounts (handle nulls).
                double amtA = a[1] is double aa ? aa : (a[1] is int ai ? ai : 0.0);
                double amtB = b[1] is double ba ? ba : (b[1] is int bi ? bi : 0.0);
                // Descending order for amount.
                return amtB.CompareTo(amtA);
            });

            // Write sorted data back to the worksheet.
            int writeRow = startRow + 1;
            foreach (var rowValues in rows)
            {
                for (int c = 0; c < totalColumns; c++)
                {
                    sheet.Cells[writeRow, startColumn + c].PutValue(rowValues[c]);
                }
                writeRow++;
            }

            // Save the sorted workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Sorted workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
