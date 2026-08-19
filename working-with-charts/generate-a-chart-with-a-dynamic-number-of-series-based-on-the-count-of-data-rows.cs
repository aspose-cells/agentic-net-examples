// Title: Create a Dynamic Column Chart with Variable Series Count Using Aspose.Cells for .NET (C#)
// Description: This example builds a new workbook, fills column A with series names and column B with values, adds a column chart, sets the category axis, then loops through each data row to add a separate series via NSeries.Add, assigning the name from column A, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells | C# chart automation | dynamic chart series | NSeries.Add | column chart from rows | Excel chart programmatically | variable series count | Aspose.Cells example
// Common Searches: add variable number of series to Aspose.Cells chart | create column chart from data rows C# | dynamic series chart Aspose.Cells .NET | loop to add chart series Aspose.Cells | Excel chart with one series per row
// Developer Intent: Generate a column chart where each worksheet row becomes an individual series using Aspose.Cells for .NET.
// Use Cases: Sales dashboard that automatically plots each product as its own series. | Performance report that expands with new metrics without code changes. | Web API that returns an Excel chart adapting to any row count in the source data.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart that adds a series for each row, using column A for names and column B for values. | Show how to iterate over worksheet rows and dynamically add series to a column chart, then set a title and save the workbook. | Provide an Aspose.Cells example that uses NSeries.Add in a loop to build a stacked column chart with a variable number of series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace DynamicSeriesChartDemo
{
    // This example builds a new workbook, fills column A with series names and column B with values, adds a column chart, sets the category axis, then loops through each data row to add a separate series via NSeries.Add, assigning the name from column A, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample data
            // Column A : Category / Series name
            // Column B : Value for each series (one value per row)
            // ------------------------------------------------------------
            sheet.Cells["A1"].PutValue("Series");
            sheet.Cells["B1"].PutValue("Value");

            // Example: 5 data rows (you can change the count to test dynamism)
            int dataRowCount = 5;
            for (int i = 0; i < dataRowCount; i++)
            {
                int rowIndex = i + 2; // data starts from row 2
                sheet.Cells[$"A{rowIndex}"].PutValue($"Series {i + 1}");
                sheet.Cells[$"B{rowIndex}"].PutValue((i + 1) * 10);
            }

            // ------------------------------------------------------------
            // Add a column chart
            // ------------------------------------------------------------
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the category (X‑axis) data – the series names in column A
            int lastDataRow = sheet.Cells.MaxDataRow; // last row that contains data
            chart.NSeries.CategoryData = $"=Sheet1!$A$2:$A${lastDataRow}";

            // ------------------------------------------------------------
            // Dynamically add a series for each data row
            // Each row has its value in column B, so we add a series that
            // points to that single cell. The series name is taken from column A.
            // ------------------------------------------------------------
            for (int row = 2; row <= lastDataRow; row++)
            {
                // Define the range for the series values (column B of the current row)
                string valueRange = $"=Sheet1!$B${row}";

                // Add the series; 'true' indicates vertical data layout
                int seriesIdx = chart.NSeries.Add(valueRange, true);

                // Assign a name to the series using the corresponding cell in column A
                chart.NSeries[seriesIdx].Name = $"=Sheet1!$A${row}";
            }

            // Optional: set a chart title
            chart.Title.Text = "Dynamic Series Chart";

            // Save the workbook
            workbook.Save("DynamicSeriesChart.xlsx");
        }
    }
}
