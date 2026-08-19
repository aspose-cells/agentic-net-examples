// Title: C# – Dynamically Update Aspose.Cells Chart Range Based on a User‑Provided Date Interval
// Description: This example creates an Excel workbook with dates in column A and values in column B, adds a column chart, then filters rows that fall between a start and end date supplied by the user. It builds a contiguous or union address string and calls SetChartDataRange and CategoryData to refresh the chart before saving the file.
// Keywords: Aspose.Cells | C# chart dynamic range | SetChartDataRange | date interval filter | runtime chart update | .NET Excel chart example | column chart data source | Excel date range selection | Aspose.Cells chart series | programmatic chart refresh
// Common Searches: Aspose.Cells change chart data range at runtime | C# filter chart series by date range | How to update Excel chart source dynamically with Aspose.Cells | SetChartDataRange example .NET | Aspose.Cells chart date interval
// Developer Intent: Adjust an existing Aspose.Cells chart so it displays only the rows whose dates lie within a user‑defined period, without recreating the chart.
// Use Cases: Display sales figures for a custom month range on a dashboard. | Create an interactive report where selecting start/end dates automatically trims the chart data. | Generate periodic Excel files that automatically show the most recent quarter’s trends.
// AI Prompts: Generate C# code that reads start and end dates from UI controls and updates an Aspose.Cells chart’s series accordingly. | Refactor the sample to encapsulate the date‑filter logic into a reusable method that returns the new range strings. | Show how to bind a chart to a named range and modify the named range definition based on a user‑selected date window.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChartDemo
{
    // This example creates an Excel workbook with dates in column A and values in column B, adds a column chart, then filters rows that fall between a start and end date supplied by the user. It builds a contiguous or union address string and calls SetChartDataRange and CategoryData to refresh the chart before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: dates in column A, values in column B
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");
            DateTime startDate = new DateTime(2023, 1, 1);
            for (int i = 0; i < 12; i++)
            {
                // One entry per month
                sheet.Cells[i + 2, 0].PutValue(startDate.AddMonths(i));
                sheet.Cells[i + 2, 1].PutValue((i + 1) * 10);
            }

            // Add a column chart that initially uses the whole data range
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set initial data range (including header row)
            chart.SetChartDataRange("B1:B13", true);
            chart.NSeries.CategoryData = "A2:A13";

            // ------------------------------------------------------------
            // USER‑SELECTED DATE RANGE (for demo purposes we hard‑code it)
            // ------------------------------------------------------------
            DateTime userStart = new DateTime(2023, 3, 1); // March 1, 2023
            DateTime userEnd   = new DateTime(2023, 8, 31); // August 31, 2023

            // Determine the rows that fall within the selected date range
            List<int> rowsInRange = new List<int>();
            for (int row = 2; row <= 13; row++) // data rows start at row 2 (index 1)
            {
                object cellValue = sheet.Cells[row - 1, 0].Value; // column A (date)
                if (cellValue is DateTime dt)
                {
                    if (dt >= userStart && dt <= userEnd)
                    {
                        rowsInRange.Add(row);
                    }
                }
            }

            // If no rows match, keep the original range
            if (rowsInRange.Count > 0)
            {
                // Build the address strings for values and categories
                string valueRange = BuildRangeString("B", rowsInRange);
                string categoryRange = BuildRangeString("A", rowsInRange);

                // Update the chart data source dynamically
                chart.SetChartDataRange(valueRange, true);
                chart.NSeries.CategoryData = categoryRange;
            }

            // Optional: display the current chart data range (for verification)
            Console.WriteLine("Current chart data range: " + chart.GetChartDataRange());

            // Save the workbook
            workbook.Save("DynamicChartByDateRange.xlsx");
        }

        // Helper method to build a contiguous range string like "B3:B8"
        // If rows are non‑contiguous, it builds a union range "B3,B5,B7"
        private static string BuildRangeString(string columnLetter, List<int> rows)
        {
            if (rows.Count == 0) return string.Empty;

            // Check if rows form a continuous block
            bool isContinuous = true;
            for (int i = 1; i < rows.Count; i++)
            {
                if (rows[i] != rows[i - 1] + 1)
                {
                    isContinuous = false;
                    break;
                }
            }

            if (isContinuous)
            {
                // Continuous block: use start:end notation
                return $"{columnLetter}{rows[0]}:{columnLetter}{rows[rows.Count - 1]}";
            }
            else
            {
                // Non‑continuous: join individual cells with commas
                List<string> parts = new List<string>();
                foreach (int r in rows)
                {
                    parts.Add($"{columnLetter}{r}");
                }
                return string.Join(",", parts);
            }
        }
    }
}
