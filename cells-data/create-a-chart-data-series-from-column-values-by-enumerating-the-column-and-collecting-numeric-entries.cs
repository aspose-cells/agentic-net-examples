// Title: Aspose.Cells C# – Build a Column Chart Series from Numeric Cells in a Column
// Description: Demonstrates how to enumerate a worksheet column, filter numeric cells, construct a non‑contiguous A1‑style range string, and assign it to a column chart series using Aspose.Cells for .NET. Includes optional category data and workbook saving.
// Keywords: Aspose.Cells chart series C# | column chart from numeric cells | enumerate worksheet column Aspose | non contiguous data range Aspose.Cells | filter numeric values Excel chart | C# Aspose.Cells example
// Common Searches: Aspose.Cells create chart series from specific cells | C# enumerate column and collect numeric cell addresses | build comma separated range for Aspose.Cells chart | add column chart with only numeric data Aspose | set category data for Aspose.Cells column chart
// Developer Intent: Create a column chart that uses only the numeric entries from a mixed‑type column by dynamically building the data range.
// Use Cases: Generate a sales bar chart while ignoring text labels in the source column. | Produce a performance chart from a column that mixes comments and values, displaying only the numbers. | Build a dynamic chart where the data range updates automatically as numeric cells are added or removed.
// AI Prompts: Write C# code with Aspose.Cells that scans a column, gathers addresses of numeric cells, and adds them as a series to a column chart. | Show how to create a comma‑separated A1‑style range from non‑contiguous numeric cells for an Aspose.Cells chart series. | Explain setting category data for a chart when the series data is assembled from filtered numeric cells.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSeriesFromColumn
{
    // Demonstrates how to enumerate a worksheet column, filter numeric cells, construct a non‑contiguous A1‑style range string, and assign it to a column chart series using Aspose.Cells for .NET. Includes optional category data and workbook saving.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate column A with mixed data (strings and numbers)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Text");
            sheet.Cells["A4"].PutValue(25);
            sheet.Cells["A5"].PutValue(30);
            sheet.Cells["A6"].PutValue("Another");
            sheet.Cells["A7"].PutValue(45);
            sheet.Cells["A8"].PutValue(0);
            sheet.Cells["A9"].PutValue("End");
            sheet.Cells["A10"].PutValue(60);

            // Enumerate the column and collect addresses of numeric cells
            List<string> numericCellAddresses = new List<string>();
            for (int row = 1; row <= 9; row++) // rows 2..10 (0‑based index)
            {
                Cell cell = sheet.Cells[row, 0]; // column A
                if (cell.Type == CellValueType.IsNumeric)
                {
                    // Build address in A1 style (e.g., $A$2)
                    string address = $"${cell.Name}";
                    numericCellAddresses.Add(address);
                }
            }

            // Build the data range string for the series (comma‑separated list)
            // Example: =Sheet1!$A$2,$A$4,$A$5,...
            string dataRange = $"=Sheet1!{string.Join(",", numericCellAddresses)}";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series using the collected numeric range (vertical orientation)
            chart.NSeries.Add(dataRange, true);

            // Optionally set category data (e.g., row numbers) for better display
            chart.NSeries.CategoryData = "=Sheet1!$B$2:$B$10";

            // Save the workbook
            workbook.Save("ChartFromNumericColumn.xlsx");
        }
    }
}
