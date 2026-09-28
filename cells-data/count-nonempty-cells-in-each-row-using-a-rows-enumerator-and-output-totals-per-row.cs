// Title: Count non‑empty cells per row using the Rows enumerator in Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over each worksheet row and prints the number of populated cells in that row. | Modify the example to skip rows that have zero non‑empty cells and display counts only for rows containing data. | Add logic to store each row’s non‑empty cell count in a Dictionary<int,int> for later analysis. | Extend the program to write the per‑row cell counts to a CSV file alongside the workbook.
// Common Searches: Aspose.Cells C# count non‑empty cells in each worksheet row | how to enumerate rows and cells with Aspose.Cells to get row wise cell totals | skip empty rows when counting cells using Aspose.Cells in .NET | store per‑row cell counts from Aspose.Cells into a collection | export Aspose.Cells row cell count results to CSV in C#
// Tags: Aspose.Cells rows enumerator non‑empty cell count | C# per‑row cell count Aspose.Cells | skip empty rows Aspose.Cells | collect row cell counts dictionary Aspose.Cells | export row counts to CSV Aspose.Cells

using System;
using System.Collections;
using Aspose.Cells;

namespace AsposeCellsRowCellCount
{
    // The sample creates a workbook, adds data to several rows, then uses the worksheet's Rows enumerator to walk each existing row, counts cells whose Value is not null and whose StringValue is not empty, prints the count per row, and finally saves the workbook as RowCellCounts.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data (some rows have empty cells)
            worksheet.Cells["A1"].PutValue("Header1");
            worksheet.Cells["B1"].PutValue("Header2");
            worksheet.Cells["A2"].PutValue("Row2Col1");
            worksheet.Cells["C2"].PutValue(123);
            worksheet.Cells["A3"].PutValue("Row3Col1");
            worksheet.Cells["B3"].PutValue("Row3Col2");
            worksheet.Cells["C3"].PutValue("Row3Col3");
            // Row 4 will be completely empty
            // Row 5 with a single non‑empty cell
            worksheet.Cells["B5"].PutValue("OnlyCell");

            // Get the rows enumerator from the worksheet
            IEnumerator rowEnumerator = worksheet.Cells.Rows.GetEnumerator();

            // Iterate through each row that exists in the worksheet
            while (rowEnumerator.MoveNext())
            {
                Row row = (Row)rowEnumerator.Current;
                int nonEmptyCellCount = 0;

                // Enumerate cells within the current row
                IEnumerator cellEnumerator = row.GetEnumerator();
                while (cellEnumerator.MoveNext())
                {
                    Cell cell = (Cell)cellEnumerator.Current;
                    // Count cell if it contains a non‑null, non‑empty value
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.StringValue))
                    {
                        nonEmptyCellCount++;
                    }
                }

                // Output the total of non‑empty cells for this row
                Console.WriteLine($"Row {row.Index + 1}: {nonEmptyCellCount} non‑empty cell(s)");
            }

            // Save the workbook (optional, just to demonstrate lifecycle)
            workbook.Save("RowCellCounts.xlsx");
        }
    }
}
