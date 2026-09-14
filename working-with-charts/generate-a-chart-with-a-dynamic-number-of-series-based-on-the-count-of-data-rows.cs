// Title: Generate an Excel column chart with a runtime‑determined number of series using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart in Aspose.Cells and adds series dynamically based on a variable seriesCount, using NSeries.Add with generated range strings. | Implement a C# helper method that converts a zero‑based column index to an Excel column letter for building chart range formulas in Aspose.Cells. | Show how to assign the X‑axis category range once and set each series name from its header cell while constructing the chart programmatically.
// Common Searches: how to add a variable number of series to an Aspose.Cells chart in C# | Aspose.Cells C# create column chart with dynamic series count | generate Excel range strings for chart series programmatically using Aspose.Cells | convert column index to Excel column letter C# Aspose.Cells helper method
// Tags: chart series count determined at runtime Aspose.Cells | NSeries.Add range generation Aspose.Cells | column index to Excel letter conversion C# | Excel column chart multiple data columns Aspose.Cells | X‑axis category range assignment Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DynamicSeriesChartDemo
{
    // The example creates a new workbook, fills column A with category labels and adds a configurable number of series columns with sample data. It then inserts a column chart, sets the X‑axis category range once, and iterates over the series count to build range strings, add each series via NSeries.Add, and assign series names from header cells. A helper method converts zero‑based column indexes to Excel column letters for the range formulas, and the workbook is saved as DynamicSeriesChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------
            // Populate sample data
            // Column A : Categories
            // Columns B..E : Series data (the number of series can change)
            // -----------------------------
            int startRow = 1; // zero‑based index (row 2 in Excel)
            int startCol = 0; // column A

            // Add category labels
            sheet.Cells[startRow, startCol].PutValue("Category");
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[startRow + i + 1, startCol].PutValue("Cat " + (i + 1));
            }

            // Add a variable number of series (e.g., 3 series)
            int seriesCount = 3; // this can be set dynamically
            for (int s = 0; s < seriesCount; s++)
            {
                int colIndex = startCol + 1 + s; // B, C, D, ...
                sheet.Cells[startRow, colIndex].PutValue("Series " + (s + 1));
                for (int i = 0; i < 5; i++)
                {
                    // Sample values: (s+1) * (i+1) * 10
                    sheet.Cells[startRow + i + 1, colIndex].PutValue((s + 1) * (i + 1) * 10);
                }
            }

            // -----------------------------
            // Create a chart
            // -----------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 1, 27, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Dynamic Series Chart";

            // Set category (X‑axis) data once
            string categoryRange = $"=Sheet1!${CellIndexToName(startCol)}${startRow + 2}:${CellIndexToName(startCol)}${startRow + 6}";
            chart.NSeries.CategoryData = categoryRange;

            // Add each series based on the number of data columns
            for (int s = 0; s < seriesCount; s++)
            {
                int colIndex = startCol + 1 + s;
                string seriesRange = $"=Sheet1!${CellIndexToName(colIndex)}${startRow + 2}:${CellIndexToName(colIndex)}${startRow + 6}";
                // isVertical = true because values are stored column‑wise
                chart.NSeries.Add(seriesRange, true);
                // Optionally set the series name (taken from the header cell)
                chart.NSeries[s].Name = $"=Sheet1!${CellIndexToName(colIndex)}${startRow + 1}";
            }

            // Save the workbook
            workbook.Save("DynamicSeriesChart.xlsx");
        }

        // Helper method to convert zero‑based column index to Excel column letters (A, B, …, AA, AB, …)
        private static string CellIndexToName(int columnIndex)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string name = string.Empty;
            int dividend = columnIndex + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                name = letters[modulo] + name;
                dividend = (dividend - modulo) / 26;
            }

            return name;
        }
    }
}
