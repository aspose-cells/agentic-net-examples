// Title: Count non‑empty cells per row with Aspose.Cells Rows enumerator in C#
// Description: The sample creates a workbook, adds a few values, obtains an IEnumerator for the worksheet's Rows collection, walks each row, counts cells whose Value is not null, prints the row number and non‑empty cell total, and saves the workbook.
// Keywords: Aspose.Cells | C# | Rows enumerator | count non empty cells | iterate worksheet rows | cell count per row | Aspose.Cells .NET example | enumerate rows and cells | non‑empty cell detection | Aspose.Cells sample code
// Common Searches: how to count non empty cells per row using Aspose.Cells | Aspose.Cells C# iterate rows and cells example | Rows.GetEnumerator count cells Aspose.Cells | C# Aspose.Cells count populated cells in each row | Aspose.Cells .NET row wise cell count
// Developer Intent: Show how to enumerate rows with Aspose.Cells and compute the number of populated cells in each row.
// Use Cases: Validate data completeness by summarizing filled cells per row. | Identify sparsely populated rows for data quality checks. | Generate row‑level statistics for storage or performance analysis.
// AI Prompts: Write C# code that uses Aspose.Cells to count non‑empty cells per row and writes the results to a new worksheet. | Explain how to modify the row‑enumeration loop to skip rows that contain zero populated cells. | Adapt the example to count non‑empty cells per column instead of per row using Aspose.Cells.

using System;
using System.Collections;
using Aspose.Cells;

// The sample creates a workbook, adds a few values, obtains an IEnumerator for the worksheet's Rows collection, walks each row, counts cells whose Value is not null, prints the row number and non‑empty cell total, and saves the workbook.
class CountNonEmptyCellsPerRow
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data
        cells["A1"].PutValue("Header");
        cells["B1"].PutValue("Data1");
        cells["C1"].PutValue(123);
        cells["A2"].PutValue("Row2Col1");
        cells["C2"].PutValue(456);
        cells["E3"].PutValue("Only cell in row 3");

        // Obtain an enumerator for the rows collection
        IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator();

        // Iterate through each row that contains data
        while (rowEnumerator.MoveNext())
        {
            Row row = (Row)rowEnumerator.Current;
            int nonEmptyCellCount = 0;

            // Enumerate cells within the current row
            IEnumerator cellEnumerator = row.GetEnumerator();
            while (cellEnumerator.MoveNext())
            {
                Cell cell = (Cell)cellEnumerator.Current;
                // Count cell if it holds a non‑null value
                if (cell != null && cell.Value != null)
                {
                    nonEmptyCellCount++;
                }
            }

            // Output the count for the current row (row index is zero‑based)
            Console.WriteLine($"Row {row.Index + 1}: {nonEmptyCellCount} non‑empty cell(s)");
        }

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("CountNonEmptyCellsPerRow.xlsx");
    }
}
