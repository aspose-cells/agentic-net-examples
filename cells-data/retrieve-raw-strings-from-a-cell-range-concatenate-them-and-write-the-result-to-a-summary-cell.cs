// Title: Aspose.Cells for .NET – Concatenate raw StringValues from a range into a summary cell
// Description: This C# example creates a workbook, fills range B2:D4 with text, numbers and dates, iterates the range using Aspose.Cells.Range, concatenates each cell's raw StringValue with a space, trims the result and writes it to cell A1 before saving as SummaryResult.xlsx.
// Keywords: Aspose.Cells concatenate range | Cell.StringValue .NET | read raw string from cells | summary cell Aspose.Cells | C# Excel string concatenation | Aspose.Cells range iteration
// Common Searches: how to join cell values into one cell using Aspose.Cells | retrieve raw string from Excel range Aspose.Cells .NET | concatenate text numbers dates Aspose.Cells example | Aspose.Cells write summary string to a cell
// Developer Intent: Read the raw string representation of every cell in a specified range, combine them into a single text string, and store that string in a designated summary cell.
// Use Cases: Create a one‑line report that merges product names, IDs and dates from a table. | Generate a searchable keyword list by joining mixed data types from a selected area. | Build a consolidated comment field for export by concatenating non‑empty cells in a user‑defined range.
// AI Prompts: Write C# code with Aspose.Cells that concatenates raw StringValues from a dynamic range, skips empty cells, and outputs to a target cell. | Show how to change the delimiter to a comma and ignore whitespace‑only cells in the concatenation example. | Explain the difference between Cell.StringValue and Cell.Value in Aspose.Cells and advise when each should be used for building text strings.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSummaryDemo
{
    // This C# example creates a workbook, fills range B2:D4 with text, numbers and dates, iterates the range using Aspose.Cells.Range, concatenates each cell's raw StringValue with a space, trims the result and writes it to cell A1 before saving as SummaryResult.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data in the range B2:D4
                cells["B2"].PutValue("Apple");
                cells["C2"].PutValue(123);               // numeric value will be converted to string
                cells["D2"].PutValue(DateTime.Today);    // date will be formatted as string
                cells["B3"].PutValue("Banana");
                cells["C3"].PutValue("Cherry");
                cells["D3"].PutValue("Date");
                cells["B4"].PutValue("Elderberry");
                cells["C4"].PutValue("Fig");
                cells["D4"].PutValue("Grape");

                // Define the range from which to gather raw string values
                string rangeAddress = "B2:D4";
                AsposeRange range = cells.CreateRange(rangeAddress);

                // Concatenate the raw string values of each cell in the range
                string concatenated = string.Empty;
                foreach (Cell cell in range)
                {
                    // StringValue returns the formatted string representation of the cell's content
                    concatenated += cell.StringValue + " ";
                }

                // Trim the trailing separator
                concatenated = concatenated.TrimEnd();

                // Write the concatenated result to a summary cell (e.g., A1)
                cells["A1"].PutValue(concatenated);

                // Save the workbook
                workbook.Save("SummaryResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
