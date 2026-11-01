// Title: Generate a column chart in Aspose.Cells for .NET by extracting only numeric values from a mixed worksheet column
// AI Prompts: Write C# code that reads a specific column, filters out non‑numeric cells, writes the numeric values to another column, and adds that range as a series to a column chart using Aspose.Cells. | Show how to build a dynamic data range for an Aspose.Cells column chart based on the count of numeric entries in a source column.
// Common Searches: how to create a column chart in Aspose.Cells using only numeric cells from a column | filter non‑numeric Excel cells before adding to chart with Aspose.Cells C# | Aspose.Cells dynamic chart series range based on numeric values in column A | C# extract numeric values from worksheet column and plot with Aspose.Cells chart | ignore text cells when building chart data series in Aspose.Cells .NET
// Tags: Aspose.Cells column chart from numeric cells | filter numeric values worksheet column Aspose.Cells | dynamic chart data range Aspose.Cells .NET | extract numeric entries Excel using Aspose.Cells | create chart series from worksheet column C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates column A with mixed data, collects only numeric cells, writes them to column B, adds that range as a series to a column chart, sets a chart title, and saves the file as ChartFromColumn.xlsx.
class ChartSeriesFromColumn
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with mixed data (numeric and non‑numeric)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue("Text");
            sheet.Cells["A3"].PutValue(25.5);
            sheet.Cells["A4"].PutValue(40);
            sheet.Cells["A5"].PutValue("Another");
            sheet.Cells["A6"].PutValue(15);
            sheet.Cells["A7"].PutValue(null);
            sheet.Cells["A8"].PutValue(30);
            sheet.Cells["A9"].PutValue("100"); // string that looks numeric, will be ignored
            sheet.Cells["A10"].PutValue(5);

            // Enumerate column A and collect only numeric entries
            List<double> numericValues = new List<double>();
            int lastRow = sheet.Cells.MaxDataRow; // last row with any data in the sheet
            for (int row = 0; row <= lastRow; row++)
            {
                Cell cell = sheet.Cells[row, 0]; // column 0 = A
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericValues.Add(cell.DoubleValue);
                }
            }

            // Write numeric values to column B (used as the series data source)
            for (int i = 0; i < numericValues.Count; i++)
            {
                sheet.Cells[i, 1].PutValue(numericValues[i]); // B column
            }

            // Write simple index labels to column C (optional, not used for categories here)
            for (int i = 0; i < numericValues.Count; i++)
            {
                sheet.Cells[i, 2].PutValue(i + 1); // C column
            }

            // Add a column chart to the worksheet
            // Parameters: chart type, upper‑left row, upper‑left column, lower‑right row, lower‑right column
            int chartUpperRow = 12;
            int chartUpperColumn = 0;
            int chartLowerRow = 30;
            int chartLowerColumn = 7;
            int chartIndex = sheet.Charts.Add(ChartType.Column, chartUpperRow, chartUpperColumn, chartLowerRow, chartLowerColumn);
            Chart chart = sheet.Charts[chartIndex];

            // Add the collected numeric values as a series (vertical orientation)
            string seriesRange = $"B1:B{numericValues.Count}";
            chart.NSeries.Add(seriesRange, true);

            // Set a title for the chart
            chart.Title.Text = "Numeric Values from Column A";

            // Save the workbook to a file
            workbook.Save("ChartFromColumn.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
