// Title: Convert a column chart to a line chart and freeze its data rows with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change chart type from column to line | How to freeze worksheet rows up to the chart's data series using Aspose.Cells | Parse chart series range to get max row index in Aspose.Cells C# | Apply FreezePanes based on chart data range with Aspose.Cells for .NET | Programmatically convert chart type and lock rows in an Excel file using Aspose.Cells
// Tags: Aspose.Cells change chart type line | Aspose.Cells freeze panes by chart data rows | C# chart series range max row Aspose.Cells | Aspose.Cells FreezePanes worksheet rows | Aspose.Cells convert column chart to line chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartManipulation
{
    // The example creates a workbook, fills columns A and B with sample data, adds a column chart, then switches the chart to a line type. It parses the chart's series and category ranges to find the highest row used, and uses FreezePanes to lock all rows up to that row. The workbook is saved as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to output file
                const string outputPath = "ChartLineAndFreezeRows.xlsx";

                // 1. Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // 2. Populate sample data (Category in column A, Values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                for (int i = 2; i <= 10; i++)
                {
                    sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                    sheet.Cells[$"B{i}"].PutValue(i * 10);
                }

                // 3. Add a chart (initially Column type) that uses the data range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 12, 0, 30, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart (both series and categories)
                chart.NSeries.Add("B2:B10", true);
                chart.NSeries.CategoryData = "A2:A10";

                // 4. Change the chart type to a Line chart
                chart.Type = ChartType.Line;

                // 5. Determine the rows that contain the chart's data series
                int maxDataRow = -1; // zero‑based index

                // Helper local function to parse a range like "B2:B10" and return the highest row index
                int GetMaxRowFromRange(string range)
                {
                    if (string.IsNullOrEmpty(range))
                        return -1;

                    // Remove any leading '=' that may be present
                    if (range.StartsWith("="))
                        range = range.Substring(1);

                    // Split start and end cells
                    string[] parts = range.Split(':');
                    if (parts.Length == 0)
                        return -1;

                    // Convert cell names to row/column indices
                    int startRow, startCol, endRow, endCol;
                    CellsHelper.CellNameToIndex(parts[0], out startRow, out startCol);
                    if (parts.Length == 2)
                    {
                        CellsHelper.CellNameToIndex(parts[1], out endRow, out endCol);
                    }
                    else
                    {
                        // Single‑cell range
                        endRow = startRow;
                    }

                    return Math.Max(startRow, endRow);
                }

                // Process each series in the chart
                foreach (Series series in chart.NSeries)
                {
                    int seriesMaxRow = GetMaxRowFromRange(series.Values);
                    if (seriesMaxRow > maxDataRow)
                        maxDataRow = seriesMaxRow;
                }

                // Also consider the CategoryData range (X‑axis labels)
                int categoryMaxRow = GetMaxRowFromRange(chart.NSeries.CategoryData);
                if (categoryMaxRow > maxDataRow)
                    maxDataRow = categoryMaxRow;

                // Freeze rows up to and including the last data row, if any rows were found
                if (maxDataRow >= 0)
                {
                    // FreezePanes expects the row/column where the split occurs and the number of rows/columns to freeze.
                    int freezeRow = maxDataRow + 1; // zero‑based index of the first unfrozen row
                    sheet.FreezePanes(freezeRow, 0, freezeRow, 0);
                    Console.WriteLine($"Frozen rows 0 through {maxDataRow} (total {freezeRow} rows).");
                }
                else
                {
                    Console.WriteLine("No data rows detected; skipping freeze operation.");
                }

                // 6. Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
