// Title: Use GetStringValue with WithFormatting to extract and sort dates in Aspose.Cells for .NET
// Description: Creates a workbook, writes mixed‑format date strings to column A, reads each cell's displayed value via GetStringValue (WithFormatting), parses the strings into DateTime objects, orders them chronologically, and writes the sorted dates to column B with a consistent built‑in date style.
// Keywords: Aspose.Cells GetStringValue WithFormatting | read formatted date string .NET | parse Excel dates C# | sort dates Aspose.Cells | Excel date style Aspose | DateTime conversion Aspose.Cells | C# Excel date sorting example
// Common Searches: GetStringValue with formatting Aspose.Cells example | how to read displayed date value from Excel using Aspose | parse and sort mixed format dates in C# with Aspose.Cells | extract formatted cell value Aspose.Cells .NET | sort Excel dates programmatically Aspose
// Developer Intent: Read the visible date text from worksheet cells, convert it to DateTime, and reorder the rows based on chronological order.
// Use Cases: Normalize a column of heterogeneous date strings before analysis. | Generate a chronologically ordered report by writing sorted dates to a new column. | Log rows that contain unparseable date strings for data‑quality review.
// AI Prompts: Write C# code that uses Aspose.Cells GetStringValue with WithFormatting to read date strings, convert them to DateTime, sort them, and output the sorted list to another column. | Show how to handle parsing errors when extracting formatted dates from Excel cells with Aspose.Cells. | Explain how to apply a built‑in date number format to both source and sorted cells in an Aspose.Cells workbook.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsDateSortingDemo
{
    // Creates a workbook, writes mixed‑format date strings to column A, reads each cell's displayed value via GetStringValue (WithFormatting), parses the strings into DateTime objects, orders them chronologically, and writes the sorted dates to column B with a consistent built‑in date style.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate column A with date strings in various formats
                string[] dateStrings = {
                    "15/05/2023",      // dd/MM/yyyy
                    "2023-05-14",      // yyyy-MM-dd
                    "May 13, 2023",    // MMM dd, yyyy
                    "2023.05.12",      // yyyy.MM.dd
                    "2023/05/11"       // yyyy/MM/dd
                };

                for (int i = 0; i < dateStrings.Length; i++)
                {
                    // Put the raw string into the cell
                    cells[i, 0].PutValue(dateStrings[i]);

                    // Apply a date style so the cell displays a date format
                    Style style = cells[i, 0].GetStyle();
                    style.Number = 14; // Built‑in date format (e.g., "m/d/yyyy")
                    cells[i, 0].SetStyle(style);
                }

                // Extract the formatted date strings using the displayed value
                List<(int Row, DateTime Date)> extractedDates = new List<(int, DateTime)>();
                for (int i = 0; i < dateStrings.Length; i++)
                {
                    // StringValue returns the displayed (formatted) value of the cell
                    string formatted = cells[i, 0].StringValue;

                    // Try to parse the string into a DateTime object using flexible parsing
                    if (DateTime.TryParse(formatted, out DateTime dt))
                    {
                        extractedDates.Add((i, dt));
                    }
                    else
                    {
                        Console.WriteLine($"Unable to parse date string at row {i + 1}: '{formatted}'");
                    }
                }

                // Sort the list by the DateTime value
                extractedDates.Sort((a, b) => a.Date.CompareTo(b.Date));

                // Write the sorted dates (as DateTime values) into column B
                for (int i = 0; i < extractedDates.Count; i++)
                {
                    DateTime sortedDate = extractedDates[i].Date;

                    // Place the DateTime value in column B (index 1)
                    cells[i, 1].PutValue(sortedDate);

                    // Apply the same date style for consistency
                    Style style = cells[i, 1].GetStyle();
                    style.Number = 14; // Built‑in date format
                    cells[i, 1].SetStyle(style);
                }

                // Save the workbook to demonstrate the result
                workbook.Save("SortedDates.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
