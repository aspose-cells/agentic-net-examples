// Title: Generate a column chart with a variable number of series from data rows using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that reads each data row and adds it as a separate series to a column chart. | Implement a helper that converts zero‑based row and column indices to Excel cell references for dynamic series ranges. | Customize the chart by setting a title and saving the workbook as an .xlsx file with Aspose.Cells. | Loop through worksheet rows to automatically determine the range for each series and assign the series name from the first column.
// Common Searches: Aspose.Cells C# create column chart where each row is a series | how to add dynamic series to Excel chart using Aspose.Cells .NET | convert zero based row column to Excel cell address Aspose.Cells helper method | set chart title and save workbook as xlsx with Aspose.Cells C# | determine used range for chart series programmatically Aspose.Cells
// Tags: dynamic series column chart Aspose.Cells C# | add chart series per worksheet row Aspose.Cells | cell index to Excel address conversion Aspose.Cells | set chart title Aspose.Cells | save workbook as xlsx Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Demonstrates creating a workbook, populating sample data, and adding a column chart where each data row is added as a separate series using Aspose.Cells for .NET, with a helper method that converts zero‑based indices to Excel cell names and optional chart title customization.
class DynamicChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Column A: Category (e.g., Item)
            // Columns B onward: Values for each series (one series per row)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["C1"].PutValue("Value2");
            sheet.Cells["D1"].PutValue("Value3");

            // Sample rows (each row will become a separate series)
            sheet.Cells["A2"].PutValue("Item 1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(20);
            sheet.Cells["D2"].PutValue(30);

            sheet.Cells["A3"].PutValue("Item 2");
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["D3"].PutValue(35);

            sheet.Cells["A4"].PutValue("Item 3");
            sheet.Cells["B4"].PutValue(20);
            sheet.Cells["C4"].PutValue(30);
            sheet.Cells["D4"].PutValue(40);

            // Determine the used range
            int firstDataRow = 1; // zero‑based index (row 2 in Excel)
            int lastDataRow = sheet.Cells.MaxDataRow; // last row with data
            int firstDataCol = 1; // column B (zero‑based)
            int lastDataCol = sheet.Cells.MaxDataColumn; // last column with data

            // Add a chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) data – using the Category column (A)
            chart.NSeries.CategoryData = $"Sheet1!A{firstDataRow + 1}:A{lastDataRow + 1}";

            // Dynamically add a series for each data row
            for (int row = firstDataRow; row <= lastDataRow; row++)
            {
                // Build the range string for the current row's values (e.g., B2:D2)
                string seriesRange = $"Sheet1!{CellIndexToName(firstDataCol, row)}:{CellIndexToName(lastDataCol, row)}";

                // Add the series; the second parameter 'true' indicates that the series uses the category data set above
                int seriesIdx = chart.NSeries.Add(seriesRange, true);

                // Set the series name from the Category column (A)
                chart.NSeries[seriesIdx].Name = sheet.Cells[row, 0].StringValue;
            }

            // Optional: set chart title
            chart.Title.Text = "Dynamic Series Chart";

            // Save the workbook
            workbook.Save("DynamicSeriesChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Helper method to convert zero‑based row/column indices to Excel cell name (e.g., (1,2) -> B3)
    static string CellIndexToName(int colIndex, int rowIndex)
    {
        // Convert column index to letters
        int dividend = colIndex + 1;
        string columnName = string.Empty;
        while (dividend > 0)
        {
            int modulo = (dividend - 1) % 26;
            columnName = Convert.ToChar(65 + modulo) + columnName;
            dividend = (dividend - modulo) / 26;
        }
        // Excel rows are 1‑based
        int excelRow = rowIndex + 1;
        return $"{columnName}{excelRow}";
    }
}
