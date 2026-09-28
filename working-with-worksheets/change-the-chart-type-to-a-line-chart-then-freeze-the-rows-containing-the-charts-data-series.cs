// Title: Convert a worksheet chart to a line chart and freeze rows containing its data series with Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to change the first chart on a worksheet to a line chart and then lock the rows that supply the chart’s series data. | Write C# code that parses chart series ranges, finds the highest row index, and applies FreezePanes to freeze all rows up to that point using Aspose.Cells.
// Common Searches: how to change a chart to line type and freeze its data rows using Aspose.Cells C# | Aspose.Cells freeze panes based on chart series range in .NET | C# program to convert Excel chart to line chart and lock rows with Aspose.Cells | extract row numbers from chart series values Aspose.Cells C# | freeze rows up to max data row of a chart in Aspose.Cells workbook
// Tags: Aspose.Cells chart type conversion | Aspose.Cells FreezePanes based on data rows | C# extract rows from chart series | Aspose.Cells worksheet chart manipulation | C# Excel chart data range parsing

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, verifies a chart exists, changes the first chart to a line chart, parses the series value and X‑value ranges to collect all involved row indices, determines the maximum row, freezes all rows up to that index with FreezePanes, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the worksheet.");
                return;
            }

            // Work with the first chart
            Chart chart = sheet.Charts[0];

            // Change chart type to Line
            chart.Type = ChartType.Line;

            // Collect row indices used by the chart's data series
            HashSet<int> dataRows = new HashSet<int>();

            foreach (var series in chart.NSeries)
            {
                if (!string.IsNullOrEmpty(series.Values))
                    AddRowsFromRange(series.Values, dataRows);

                if (!string.IsNullOrEmpty(series.XValues))
                    AddRowsFromRange(series.XValues, dataRows);
            }

            // Freeze rows up to the highest data row (if any)
            if (dataRows.Count > 0)
            {
                int maxRowIndex = -1; // zero‑based
                foreach (int rowIdx in dataRows)
                    if (rowIdx > maxRowIndex) maxRowIndex = rowIdx;

                // Freeze rows: use overload with four parameters (row, column, totalRows, totalColumns)
                sheet.FreezePanes(maxRowIndex + 1, 0, maxRowIndex + 1, 0);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Parses a range address (e.g., "A2:B10") and adds all row indices (zero‑based) to the set
    static void AddRowsFromRange(string rangeAddress, HashSet<int> rows)
    {
        // Split multi‑area ranges (e.g., "A2:B5,C7:D9")
        string[] areas = rangeAddress.Split(',');

        foreach (string area in areas)
        {
            // Determine start and end cells
            string[] cells = area.Split(':');
            string startCell = cells[0];
            string endCell = cells.Length > 1 ? cells[1] : startCell;

            int startRow = GetRowIndex(startCell);
            int endRow = GetRowIndex(endCell);

            // Add every row between startRow and endRow (inclusive)
            for (int r = startRow; r <= endRow; r++)
                rows.Add(r);
        }
    }

    // Converts a cell address like "B12" to a zero‑based row index (12 -> 11)
    static int GetRowIndex(string cellAddress)
    {
        // Remove column letters, keep the numeric part
        string rowPart = Regex.Replace(cellAddress.ToUpper(), "[A-Z]+", "");
        if (int.TryParse(rowPart, out int rowNumber))
            return rowNumber - 1; // zero‑based
        throw new ArgumentException($"Invalid cell address: {cellAddress}");
    }
}
