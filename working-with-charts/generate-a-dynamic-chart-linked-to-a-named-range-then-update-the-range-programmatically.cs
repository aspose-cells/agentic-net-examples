// Title: Create a dynamic column chart linked to a named range and expand it programmatically using Aspose.Cells for .NET (C#)
// Common Searches: how to bind a column chart to a named range with Aspose.Cells in C# | c# expand a named range and refresh the chart using Aspose.Cells | add new rows to Excel data and update chart source programmatically Aspose.Cells | dynamic chart data source update example with Aspose.Cells C# | programmatically modify chart series after extending data range Aspose.Cells
// Tags: Aspose.Cells column chart named range | programmatic named range expansion Aspose.Cells | chart series refresh Aspose.Cells C# | dynamic Excel chart data source Aspose.Cells | C# update chart after adding rows Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a workbook, defines a named range covering A1:B5, adds a column chart that uses this range for values and categories, saves the file, then appends June and July rows, expands the named range to A1:B7, updates the chart series to reference the new range, and saves the updated workbook.
class DynamicChartExample
{
    static void Main()
    {
        try
        {
            // Initialize a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // Populate initial data (A1:B5)
            // Column A: Categories, Column B: Values
            string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
            double[] values = { 120, 150, 180, 130, 170 };

            for (int i = 0; i < categories.Length; i++)
            {
                sheet.Cells[i, 0].PutValue(categories[i]); // A column
                sheet.Cells[i, 1].PutValue(values[i]);    // B column
            }

            // Define a named range "MyData" that refers to A1:B5
            int firstRow = 0, firstColumn = 0;
            int totalRows = categories.Length;
            int totalColumns = 2;
            int nameIndex = workbook.Worksheets.Names.Add("MyData");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            string startCell = $"{CellsHelper.ColumnIndexToName(firstColumn)}{firstRow + 1}";
            string endCell = $"{CellsHelper.ColumnIndexToName(firstColumn + totalColumns - 1)}{firstRow + totalRows}";
            namedRange.RefersTo = $"=DataSheet!${startCell}:${endCell}";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Monthly Sales";

            // Set the chart's data source to the named range "MyData"
            chart.NSeries.Add("=MyData!$B$1:$B$5", true);
            chart.NSeries[0].XValues = "=MyData!$A$1:$A$5";

            // Save the initial workbook with the chart
            workbook.Save("DynamicChart_Initial.xlsx");

            // ------------------------------
            // Programmatically update the data range
            // ------------------------------

            // Append new data rows (June and July)
            string[] newCategories = { "Jun", "Jul" };
            double[] newValues = { 190, 160 };
            int startRow = categories.Length; // row index where new data starts

            for (int i = 0; i < newCategories.Length; i++)
            {
                sheet.Cells[startRow + i, 0].PutValue(newCategories[i]); // Column A
                sheet.Cells[startRow + i, 1].PutValue(newValues[i]);    // Column B
            }

            // Expand the named range "MyData" to include the new rows (now A1:B7)
            int newTotalRows = categories.Length + newCategories.Length;
            endCell = $"{CellsHelper.ColumnIndexToName(firstColumn + totalColumns - 1)}{firstRow + newTotalRows}";
            namedRange.RefersTo = $"=DataSheet!${startCell}:${endCell}";

            // Refresh the chart to reflect the updated range
            chart.NSeries[0].Values = $"=MyData!$B$1:$B${newTotalRows}";
            chart.NSeries[0].XValues = $"=MyData!$A$1:$A${newTotalRows}";

            // Save the updated workbook
            workbook.Save("DynamicChart_Updated.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
