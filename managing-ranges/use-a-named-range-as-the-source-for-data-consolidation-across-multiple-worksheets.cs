// Title: Consolidate numeric data from several worksheets using a named range with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that defines a named range (e.g., A1:C4) on each source worksheet, iterates through the range to collect headers and sum numeric cells, and writes the aggregated results to a new worksheet using Aspose.Cells. | Show how to create a Range object from a named range, aggregate values across multiple sheets, and save the workbook as an .xlsx file with Aspose.Cells. | Provide a step‑by‑step example that extracts a named range from three worksheets, merges the data by summing matching cells, and produces a consolidated sheet.
// Common Searches: aspnet c# how to use a named range for consolidating data across multiple Excel sheets with Aspose.Cells | Aspose.Cells sum values from the same range in several worksheets and create a summary sheet | C# example consolidating data from multiple worksheets using Aspose.Cells range object | merge numeric tables from different sheets into one sheet using Aspose.Cells .NET
// Tags: Aspose.Cells consolidate worksheets using named range | C# aggregate range values across multiple sheets | Aspose.Cells create summary worksheet from multiple sources | Aspose.Cells .xlsx data aggregation with range object | C# sum numeric cells across worksheets Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The sample creates a workbook with three worksheets, each containing numeric data in the A1:C4 area. It defines a range on each sheet, captures row and column headers, sums the numeric cells across all worksheets, writes the combined headers and totals to a new "Consolidated" worksheet, and saves the result as ConsolidatedResult.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Add three worksheets with sample data
            for (int i = 0; i < 3; i++)
            {
                Worksheet ws = wb.Worksheets[i];
                ws.Name = $"Sheet{i + 1}";

                // Populate sample data in range A1:C4
                for (int row = 0; row < 4; row++)
                {
                    ws.Cells[row, 0].PutValue(row + 1 + i * 10);               // Column A (row header)
                    ws.Cells[row, 1].PutValue((row + 1) * 2 + i * 10);         // Column B
                    ws.Cells[row, 2].PutValue((row + 1) * 3 + i * 10);         // Column C
                }
            }

            // Add a destination worksheet for consolidation
            Worksheet dest = wb.Worksheets.Add("Consolidated");

            // Prepare structures to hold headers and summed values
            List<string> rowHeaders = new List<string>();
            List<string> colHeaders = new List<string>();
            Dictionary<(int row, int col), double> sums = new Dictionary<(int, int), double>();

            // Process each source worksheet
            for (int i = 0; i < 3; i++)
            {
                Worksheet src = wb.Worksheets[i];

                // Retrieve the data range (A1:C4) – use fully qualified Aspose.Cells.Range to avoid ambiguity
                Aspose.Cells.Range srcRange = src.Cells.CreateRange("A1:C4");

                for (int r = 0; r < srcRange.RowCount; r++)
                {
                    for (int c = 0; c < srcRange.ColumnCount; c++)
                    {
                        // Skip the top‑left cell (intersection of headers)
                        if (r == 0 && c == 0) continue;

                        // Capture column headers (first row, excluding top‑left)
                        if (r == 0)
                        {
                            string colHeader = srcRange[r, c].StringValue;
                            while (colHeaders.Count <= c - 1) colHeaders.Add(string.Empty);
                            colHeaders[c - 1] = colHeader;
                            continue;
                        }

                        // Capture row headers (first column, excluding top‑left)
                        if (c == 0)
                        {
                            string rowHeader = srcRange[r, c].StringValue;
                            while (rowHeaders.Count <= r - 1) rowHeaders.Add(string.Empty);
                            rowHeaders[r - 1] = rowHeader;
                            continue;
                        }

                        // Sum numeric data (excluding headers)
                        double val = srcRange[r, c].DoubleValue;
                        var key = (row: r - 1, col: c - 1);
                        if (sums.ContainsKey(key))
                            sums[key] += val;
                        else
                            sums[key] = val;
                    }
                }
            }

            // Write headers to the destination sheet
            dest.Cells[0, 0].PutValue(string.Empty); // top‑left corner
            for (int c = 0; c < colHeaders.Count; c++)
                dest.Cells[0, c + 1].PutValue(colHeaders[c]);

            for (int r = 0; r < rowHeaders.Count; r++)
                dest.Cells[r + 1, 0].PutValue(rowHeaders[r]);

            // Write summed values
            foreach (var kvp in sums)
            {
                int destRow = kvp.Key.row + 1; // offset for header row
                int destCol = kvp.Key.col + 1; // offset for header column
                dest.Cells[destRow, destCol].PutValue(kvp.Value);
            }

            // Save the workbook
            wb.Save("ConsolidatedResult.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
