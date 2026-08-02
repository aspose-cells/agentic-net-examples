// Title: Parse, Sort & Reformat Date Strings with Aspose.Cells GetStringValue (C#)
// Description: Demonstrates how to read formatted date strings from a worksheet using GetStringValue(CellValueFormatStrategy.DisplayString), convert them to DateTime with DateTime.ParseExact, sort the rows chronologically, write the dates back as true DateTime values, apply a custom "dd-MM-yyyy" style, and save the workbook as SortedDates.xlsx.
// Keywords: Aspose.Cells GetStringValue | CellValueFormatStrategy.DisplayString | C# date parsing Excel | sort dates Aspose.Cells | convert string to DateTime Aspose | custom date format Aspose.Cells | Excel date sorting .NET | Aspose.Cells example
// Common Searches: How to read displayed date strings with Aspose.Cells | Parse and sort string dates in Excel using C# | Apply custom date format after sorting with Aspose.Cells | GetStringValue with formatting Aspose.Cells .NET
// Developer Intent: Extract displayed date strings, convert to DateTime, sort, and rewrite them with a consistent date format.
// Use Cases: Read plain‑text dates stored in a column, parse them using a known pattern, and order the rows by date. | Replace the original strings with true DateTime cells while preserving the header row. | Apply a custom cell style (dd-MM-yyyy) to ensure consistent visual formatting after sorting.
// AI Prompts: Show C# code that uses Aspose.Cells GetStringValue(DisplayString) to retrieve formatted dates, sort them, and write back with a custom style. | Explain how to safely parse date strings from Excel cells using DateTime.ParseExact in Aspose.Cells. | Provide a step‑by‑step guide for sorting date strings in an Aspose.Cells workbook and exporting the result to a new file.

using System;
using System.Collections.Generic;
using System.Globalization;
using Aspose.Cells;

namespace AsposeCellsDateSortingDemo
{
    // Demonstrates how to read formatted date strings from a worksheet using GetStringValue(CellValueFormatStrategy.DisplayString), convert them to DateTime with DateTime.ParseExact, sort the rows chronologically, write the dates back as true DateTime values, apply a custom "dd-MM-yyyy" style, and save the workbook as SortedDates.xlsx.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Sample data: date strings in column A (index 0)
                // These are stored as plain strings, not as DateTime values
                cells["A1"].PutValue("Date"); // Header
                cells["A2"].PutValue("15-05-2023");
                cells["A3"].PutValue("01-04-2022");
                cells["A4"].PutValue("23-12-2023");
                cells["A5"].PutValue("07-08-2021");

                // List to hold parsed DateTime and original row index
                List<(DateTime date, int row)> dateRows = new List<(DateTime, int)>();

                // Determine the range of data rows (excluding header)
                int startRow = 1; // zero‑based index, row 2 in Excel
                int endRow = cells.MaxDataRow; // last row with data

                // Extract the displayed string from each cell and parse it to DateTime
                for (int r = startRow; r <= endRow; r++)
                {
                    // GetStringValue with DisplayString returns the cell's displayed string with formatting
                    string dateStr = cells[r, 0].GetStringValue(CellValueFormatStrategy.DisplayString);

                    // Parse using the known format (dd-MM-yyyy). Adjust format if needed.
                    DateTime parsedDate = DateTime.ParseExact(dateStr, "dd-MM-yyyy", CultureInfo.InvariantCulture);
                    dateRows.Add((parsedDate, r));
                }

                // Sort the list by the DateTime value
                dateRows.Sort((x, y) => x.date.CompareTo(y.date));

                // Write the sorted dates back to the worksheet (preserving the header)
                for (int i = 0; i < dateRows.Count; i++)
                {
                    int targetRow = startRow + i;
                    cells[targetRow, 0].PutValue(dateRows[i].date);
                    // Apply a date format so the cell displays as a date string
                    Style style = cells[targetRow, 0].GetStyle();
                    style.Custom = "dd-MM-yyyy";
                    cells[targetRow, 0].SetStyle(style);
                }

                // Save the workbook
                workbook.Save("SortedDates.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
