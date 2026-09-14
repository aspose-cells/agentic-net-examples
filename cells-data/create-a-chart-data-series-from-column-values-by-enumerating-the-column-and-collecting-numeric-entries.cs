// Title: Generate a column chart from only numeric cells in a mixed‑type column using Aspose.Cells for .NET
// AI Prompts: Write C# code that iterates through a worksheet column, selects only numeric cells, builds a comma‑separated address range, and adds it as a series to a column chart with Aspose.Cells. | Create a method that returns a range string of numeric cell addresses and demonstrates using that string to populate NSeries data in an Aspose.Cells chart. | Show how to assign the category axis to a header cell while the data series is built from a dynamically assembled numeric range in Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# create column chart using only numeric values from a column | filter non‑numeric cells when building chart series in Aspose.Cells .NET | build range string of specific cell addresses for chart series Aspose.Cells | dynamic data range for column chart from mixed data column Aspose.Cells
// Tags: numeric cell filtering for chart series Aspose.Cells | dynamic range construction for column chart C# | add column chart with selective data Aspose.Cells | mixed data column handling Aspose.Cells | set chart category to header cell Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills column A with mixed text and numbers, extracts the addresses of numeric cells, builds a comma‑separated range string, adds a column chart, assigns the numeric range as the series data, sets the category axis to the header cell, and saves the file as ChartFromNumericColumn.xlsx.
class CreateChartFromNumericColumn
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate column A with mixed data (numeric and non‑numeric)
            string[] data = { "Header", "10", "Apple", "20", "30", "Banana", "40", "50", "Cherry", "60" };
            for (int i = 0; i < data.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(data[i]);
            }

            // Collect addresses of numeric cells
            List<string> numericAddresses = new List<string>();
            for (int row = 0; row < data.Length; row++)
            {
                Cell cell = sheet.Cells[row, 0];
                if (cell.Type == CellValueType.IsNumeric)
                {
                    numericAddresses.Add(cell.Name);
                }
            }

            // Build the range string for the numeric values (e.g., Sheet1!A2,A4,A5)
            string numericRange = $"{sheet.Name}!{string.Join(",", numericAddresses)}";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add the series using the collected numeric range (vertical orientation)
            chart.NSeries.Add(numericRange, true);

            // Set category data (using the header cell)
            chart.NSeries.CategoryData = $"{sheet.Name}!A1";

            // Save the workbook
            workbook.Save("ChartFromNumericColumn.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
