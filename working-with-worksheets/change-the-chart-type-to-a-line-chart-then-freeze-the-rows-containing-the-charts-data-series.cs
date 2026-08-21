// Title: Aspose.Cells for .NET: Convert a Column Chart to a Line Chart and Freeze Its Data Rows (C#)
// Description: A concise C# example that creates a workbook, adds sample data, inserts a column chart, switches the chart type to Line, extracts the first series range, calculates the start and end rows, and freezes those rows with FreezePanes before saving the file.
// Keywords: Aspose.Cells C# | .NET Excel automation | change chart type Aspose.Cells | line chart Aspose.Cells | freeze rows Excel C# | FreezePanes Aspose.Cells | chart series range extraction | GitHub Aspose.Cells example | Excel dashboard code
// Common Searches: Aspose.Cells change column chart to line chart C# | How to freeze rows based on chart data range using Aspose.Cells | C# FreezePanes after creating a chart with Aspose.Cells | Extract chart series range Aspose.Cells .NET | Aspose.Cells example for dynamic chart type and row freezing
// Developer Intent: Update an existing chart to a line type and lock the worksheet rows that contain its data series.
// Use Cases: Financial reports that display trend lines while keeping the underlying data visible during scrolling. | Interactive Excel dashboards where chart types are switched programmatically and related rows stay in view. | Automated workbook generation that standardizes chart styles and improves navigation in large data sheets.
// AI Prompts: Generate C# code with Aspose.Cells to convert a column chart to a line chart and freeze rows up to the last row of its first data series. | Provide a reusable method that reads a chart's NSeries range, parses row numbers, and applies FreezePanes on the worksheet. | Explain how to handle multiple chart series when determining which rows to freeze in an Aspose.Cells workbook.

using System;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Charts;

// A concise C# example that creates a workbook, adds sample data, inserts a column chart, switches the chart type to Line, extracts the first series range, calculates the start and end rows, and freezes those rows with FreezePanes before saving the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data (A1:B6)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 2; i <= 6; i++)
        {
            sheet.Cells[i - 1, 0].PutValue("Item " + (i - 1));
            sheet.Cells[i - 1, 1].PutValue((i - 1) * 10);
        }

        // Add a column chart (will be changed later)
        int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];
        chart.SetChartDataRange("A1:B6", true);

        // Change the chart type to a line chart
        chart.Type = ChartType.Line;

        // Determine the rows that contain the chart's data series
        // Assuming the first series holds the data range
        string dataRange = chart.NSeries[0].Values; // e.g., "Sheet1!$B$2:$B$6"
        int startRow = 0, endRow = 0;

        if (!string.IsNullOrEmpty(dataRange))
        {
            // Remove sheet name and '$' characters
            int exclPos = dataRange.IndexOf('!');
            string rangePart = dataRange.Substring(exclPos + 1).Replace("$", ""); // e.g., "B2:B6"
            string[] parts = rangePart.Split(':');
            if (parts.Length == 2)
            {
                startRow = int.Parse(Regex.Match(parts[0], @"\d+").Value);
                endRow = int.Parse(Regex.Match(parts[1], @"\d+").Value);
            }
        }

        // Freeze rows up to the last data row (including header)
        // FreezePanes(row, column, freezedRows, freezedColumns)
        // Use 0‑based indices; row parameter is the cell where the split occurs
        if (endRow > 0)
        {
            sheet.FreezePanes(endRow, 0, endRow, 0);
        }

        // Save the workbook
        workbook.Save("ChartLineAndFreezeRows.xlsx", SaveFormat.Xlsx);
    }
}
