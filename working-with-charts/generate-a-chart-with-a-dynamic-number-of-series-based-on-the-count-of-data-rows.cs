// Title: C# – Build an Excel column chart with a dynamic number of series using Aspose.Cells
// Description: This example creates a workbook, fills columns A‑D with categories and numeric series, determines the used range, adds a column chart, assigns the X‑axis categories from column A, loops through each data column to add series automatically, sets series names from the header row, and saves the file as DynamicSeriesChart.xlsx.
// Keywords: Aspose.Cells | C# chart automation | dynamic chart series | Excel column chart programmatically | runtime series addition | set chart category data | set series names from header | Excel reporting with Aspose.Cells | chart generation .NET | used range for chart
// Common Searches: Aspose.Cells add variable number of series to chart | C# create Excel chart with dynamic columns | set series names from header row Aspose.Cells | determine used range for chart data .NET | generate column chart from worksheet data programmatically | dynamic series chart Aspose.Cells example
// Developer Intent: Create an Excel chart whose series count is derived automatically from the worksheet columns at runtime.
// Use Cases: Sales dashboard that converts each product column into its own series without hard‑coding the count. | Reporting tool that visualizes any number of metric columns from a data table with a single adaptive chart. | Export of variable‑size data sets to Excel where the chart updates automatically as new columns are added.
// AI Prompts: Show how to switch the chart to a stacked column while preserving the dynamic series logic. | Provide code to add data labels to each series that is added dynamically. | Explain how to customize the legend position and style after creating the dynamic series chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicSeriesDemo
{
    // This example creates a workbook, fills columns A‑D with categories and numeric series, determines the used range, adds a column chart, assigns the X‑axis categories from column A, loops through each data column to add series automatically, sets series names from the header row, and saves the file as DynamicSeriesChart.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data:
                // Column A – categories, Columns B..N – series values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");

                // Add some rows of data
                for (int row = 2; row <= 6; row++)
                {
                    sheet.Cells[row - 1, 0].PutValue("Item " + (row - 1)); // Category column
                    sheet.Cells[row - 1, 1].PutValue(row * 10);          // Series1 values
                    sheet.Cells[row - 1, 2].PutValue(row * 15);          // Series2 values
                    sheet.Cells[row - 1, 3].PutValue(row * 20);          // Series3 values
                }

                // Determine the used range
                int firstDataRow = 2; // 1‑based row where data starts
                int lastDataRow = sheet.Cells.MaxDataRow + 1; // MaxDataRow is zero‑based
                int firstDataColumn = 1; // column B (0‑based) – first series column
                int lastDataColumn = sheet.Cells.MaxDataColumn; // last column with data (0‑based)

                // Add a chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 8, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the category (X‑axis) data – first column (A)
                string categoryRange = $"{CellsHelper.CellIndexToName(0, 0)}{firstDataRow}:{CellsHelper.CellIndexToName(0, 0)}{lastDataRow}";
                chart.NSeries.CategoryData = categoryRange;

                // Add a series for each data column dynamically
                for (int col = firstDataColumn; col <= lastDataColumn; col++)
                {
                    // Build the range string for the current series column (no $ signs)
                    string colLetter = CellsHelper.CellIndexToName(0, col);
                    string seriesRange = $"{colLetter}{firstDataRow}:{colLetter}{lastDataRow}";

                    // Add the series (isVertical = true means column‑wise)
                    chart.NSeries.Add(seriesRange, true);
                }

                // Optional: give each series a name from the header row
                // This uses SetSeriesNames starting at index 0, reading names from row 1
                string nameRange = $"{CellsHelper.CellIndexToName(0, firstDataColumn)}1:{CellsHelper.CellIndexToName(0, lastDataColumn)}1";
                chart.NSeries.SetSeriesNames(0, nameRange, true);

                // Save the workbook
                workbook.Save("DynamicSeriesChart.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
