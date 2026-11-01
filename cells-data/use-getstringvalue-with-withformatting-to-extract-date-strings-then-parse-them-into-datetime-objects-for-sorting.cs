// Title: Read formatted dates with Cell.StringValue, convert to DateTime, and sort Excel rows using Aspose.Cells for .NET
// Common Searches: How to read a formatted date from an Excel cell using Aspose.Cells C# | Sorting worksheet rows by a date column with Aspose.Cells .NET example | Parse displayed date strings to DateTime when extracting data with Aspose.Cells | Copy rows while preserving styles after date‑based sorting in Aspose.Cells | Aspose.Cells GetStringValue versus GetValue for date parsing in C#
// Tags: Cell.StringValue date parsing Aspose.Cells | sort worksheet rows by DateTime Aspose.Cells | copy rows with styles Aspose.Cells | extract formatted dates Excel C# | Aspose.Cells date column sorting example

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The sample loads an Excel workbook, reads each cell in column A using Cell.StringValue to obtain the displayed date string, parses the string into a DateTime with the current culture, collects successful parses with their original row indices, orders the rows by the parsed dates, and copies each entire row—including values, styles, and formulas—to a new worksheet before saving the sorted workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file '{inputPath}' not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sourceSheet = workbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Assume dates are in column A (index 0) starting from row 0
            int startRow = 0;
            int endRow = sourceCells.MaxDataRow; // last row with data

            // List to hold parsed dates and their original row indices
            List<(DateTime Date, int RowIndex)> dateRows = new List<(DateTime, int)>();

            // Extract date strings with formatting and parse them
            for (int row = startRow; row <= endRow; row++)
            {
                Cell cell = sourceCells[row, 0];
                // Get the cell's displayed string (respecting formatting)
                string dateStr = cell.StringValue;

                // Try to parse the string into DateTime using the current culture
                if (DateTime.TryParse(dateStr, CultureInfo.CurrentCulture, DateTimeStyles.None, out DateTime parsedDate))
                {
                    dateRows.Add((parsedDate, row));
                }
                // If parsing fails, the row is simply skipped
            }

            // Sort the list by the parsed DateTime
            var sortedDateRows = dateRows.OrderBy(dr => dr.Date).ToList();

            // Create a new worksheet to store sorted data
            Worksheet sortedSheet = workbook.Worksheets.Add("SortedData");
            Cells sortedCells = sortedSheet.Cells;

            // Determine the maximum column index to copy all columns
            int maxColumn = sourceCells.MaxDataColumn;

            // Copy rows to the new sheet in sorted order
            int targetRow = 0;
            foreach (var (date, originalRow) in sortedDateRows)
            {
                // Copy each cell in the row (value, style, formula, etc.)
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell srcCell = sourceCells[originalRow, col];
                    Cell destCell = sortedCells[targetRow, col];

                    // Copy value
                    destCell.PutValue(srcCell.Value);

                    // Copy style
                    destCell.SetStyle(srcCell.GetStyle());

                    // Copy formula if present
                    if (!string.IsNullOrEmpty(srcCell.Formula))
                    {
                        destCell.Formula = srcCell.Formula;
                    }
                }

                targetRow++;
            }

            // Save the workbook with sorted data
            workbook.Save(outputPath);
            Console.WriteLine($"Sorted workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
