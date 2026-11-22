// Title: How to count non‑empty cells in each worksheet row with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells iterate rows and count cells that are not blank | Get number of non‑empty cells per row using Aspose.Cells Rows collection | Aspose.Cells count non‑empty cells in a row C# example | Skip null and empty string values when counting cells with Aspose.Cells | Console output row index and non‑blank cell count Aspose.Cells C#
// Tags: Rows enumerator cell count Aspose.Cells | detect empty cells Aspose.Cells C# | per‑row non‑blank cell tally Aspose.Cells | worksheet row iteration Aspose.Cells | console row index cell count Aspose.Cells

using Aspose.Cells;
using System;

// The example creates a workbook, fills sample data, then uses the worksheet's Rows enumerator to walk through each row. For every row it iterates over its cells, treats null or empty strings as empty, counts the remaining non‑empty cells, writes the row index and count to the console, and finally saves the workbook as an XLSX file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook(); // create rule

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Sample data (you can replace this with loading real data)
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["B1"].PutValue(123);
            sheet.Cells["C1"].PutValue(null);
            sheet.Cells["A2"].PutValue("");
            sheet.Cells["B2"].PutValue(45);
            sheet.Cells["C2"].PutValue("World");

            // Enumerate each row in the worksheet
            foreach (Row row in sheet.Cells.Rows)
            {
                int nonEmptyCount = 0;

                // Enumerate each cell in the current row
                foreach (Cell cell in row)
                {
                    // Determine if the cell is considered empty
                    bool isEmpty = cell.Value == null ||
                                   (cell.Value is string s && string.IsNullOrEmpty(s));

                    if (!isEmpty)
                    {
                        nonEmptyCount++;
                    }
                }

                // Output the count for the current row (row index is zero‑based)
                Console.WriteLine($"Row {row.Index}: {nonEmptyCount} non‑empty cells");
            }

            // Save the workbook if needed
            workbook.Save("output.xlsx"); // save rule
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
